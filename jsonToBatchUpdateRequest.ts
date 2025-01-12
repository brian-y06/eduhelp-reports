import fs from "fs";
const fsp = fs.promises;
import {
  changePageDimensions,
  changePageMargins,
  changePageOrientation,
} from "./pageUpdates";
import { handleParagraph } from "./paragraphUpdates";
import { handleTable } from "./tableUpdates";
import {
  findStructuralElementType,
  getStylingFieldMask,
} from "./batchUpdateHelpers";
import { docs_v1 } from "googleapis";
import { GDocRequest, InlineObject, Paragraph } from "../types/googleDocsTypes";

/**
 * This is a code sample to show my coding style and provides a high level overview of how the Google Doc to Batch Request feature is created.
 */

/**
 * A function that turns a Google Doc JSON into a batch request object for docs.documents.batchUpdate();
 * @param {JSON} docJSON the entire document JSON object that is returned from the Google Docs API .get() method on the .data property
 */
function jsonToBatchUpdateRequest(docJSON: docs_v1.Schema$Document) {
  try {
    let content = docJSON.body!.content as docs_v1.Schema$StructuralElement[];
    const inlineObjects = docJSON.inlineObjects;

    const docStyle = docJSON.documentStyle;

    // Getting the page dimensions
    const docPageSize = docStyle?.pageSize;
    const widthMagnitude = docPageSize?.width?.magnitude;
    const heightMagnitude = docPageSize?.height?.magnitude;

    const changePageSizeReq = changePageDimensions({
      height: heightMagnitude as number,
      width: widthMagnitude as number,
    });

    // Getting the margins of the page
    const marginTopMagnitude = docStyle?.marginTop?.magnitude;
    const marginBottomMagnitude = docStyle?.marginBottom?.magnitude;
    const marginRightMagnitude = docStyle?.marginRight?.magnitude;
    const marginLeftMagnitude = docStyle?.marginLeft?.magnitude;

    // Requests to change page margins
    const changePageMarginsReq = changePageMargins({
      marginTop: marginTopMagnitude as number,
      marginBottom: marginBottomMagnitude as number,
      marginRight: marginRightMagnitude as number,
      marginLeft: marginLeftMagnitude as number,
    });

    // Getting the custom styling flags like flipPageOrientation
    const flipPageOrientationFlag = docStyle?.flipPageOrientation;
    const flipOrientationReq = changePageOrientation(
      flipPageOrientationFlag as boolean,
    );

    // Under the last line of a BLUEPRINT there should be just enough space for a line with fontsize 1. This is because the document originally has a \n when it's created - and after we insert everything, the original newline will become the last line (so we will have an extra page if the BLUEPRINT goes all the way to the bottom of the page) A workaround is setting the original newline to fontsize 1 (and reducing all spacing) and leaving just a little bit of space for the newline to exists without creating a newpage.
    const shrinkOriginalNewLine = [
      {
        updateTextStyle: {
          textStyle: { fontSize: { magnitude: 1, unit: "PT" } },
          fields: "fontSize",
          range: { startIndex: 1, endIndex: 2, segmentId: "" },
        },
      },
      {
        updateParagraphStyle: {
          paragraphStyle: {
            lineSpacing: 0.06,
            spaceAbove: { magnitude: 0, unit: "PT" },
            spaceBelow: { magnitude: 0, unit: "PT" },
          },
          fields: "lineSpacing,spaceAbove,spaceBelow",
          range: { startIndex: 1, endIndex: 2, segmentId: "" },
        },
      },
    ];

    // Always run the following changes when generating a new document.
    const requests: docs_v1.Schema$Request[] = [
      ...shrinkOriginalNewLine,
      changePageSizeReq,
      flipOrientationReq,
      changePageMarginsReq,
    ];

    // Inserting a table automatically creates a newline above it, we must flag this to make sure we don't copy the parse the newline and copy it ourselves again. 
    let newLineStyleUpdatesForAfterTableInsertion = null;
    for (let i = 0; i < (content as []).length; i++) {
      const structuralElement = content[i];
      const startIndex = structuralElement?.startIndex as number;
      const endIndex = structuralElement?.endIndex as number;

      const structuralElementType =
        findStructuralElementType(structuralElement);

      switch (structuralElementType) {
        case "paragraph": {
          const {
            insertParagraphElementRequests,
            stylesForNewlineBeforeTable,
          } = handleParagraph({
            paragraphElement: structuralElement.paragraph as Paragraph,
            content,
            currentIndex: i,
            inlineObjects: inlineObjects as Record<string, InlineObject>, 
            paragraphStartIndex: startIndex as number,
            paragraphEndIndex: endIndex as number,
            isLastContentOfTableCell: false,
          });
          newLineStyleUpdatesForAfterTableInsertion =
            stylesForNewlineBeforeTable;
          requests.push(...(insertParagraphElementRequests as GDocRequest[]));

          break;
        }
          
        case "sectionBreak": {
          const sectionBreakInsertionRequest = handleSectionBreak({
            sectionBreakElement:
              structuralElement.sectionBreak as docs_v1.Schema$SectionBreak,
            sectionBreakStartIndex: startIndex,
            sectionBreakEndIndex: endIndex,
          });
          requests.push(...sectionBreakInsertionRequest);
          break;
        }
          
        case "table": {
          const tableInsertionRequest = handleTable({
            tableElement: structuralElement.table as docs_v1.Schema$Table,
            tableStartIndex: startIndex,
            tableEndIndex: endIndex,
            inlineObjects: inlineObjects as Record<string, InlineObject>,
          });
          requests.push(...tableInsertionRequest);
          if (newLineStyleUpdatesForAfterTableInsertion !== null) {
            requests.push(newLineStyleUpdatesForAfterTableInsertion);
            newLineStyleUpdatesForAfterTableInsertion = null;
          }
          break;
        }
      }
    }

    return requests;
  } catch (err) {
    if (err instanceof Error) {
      console.error(
        `There was an error in Google Doc JSON conversion to batch requests -  ${err}\n${err.stack}`,
      );
    }
  }
}

// SECTION BREAK STRUCTURAL ELEMENT HANDLING
/**
 * A method to return the appropriate type of section break requests based on whether there is a start index in the structural element or not.
 * @param {docs_v1.Schema$SectionBreak} sectionBreakElement a section break object
 * @param {number} sectionBreakStartIndex the start index of the section break
 * @param {number} sectionBreakEndIndex the end index of the section break
 * @returns an array of requests objects to insert a section break and update its style
 */
function handleSectionBreak({
  sectionBreakElement,
  sectionBreakStartIndex,
  sectionBreakEndIndex,
}: {
  sectionBreakElement: docs_v1.Schema$SectionBreak;
  sectionBreakStartIndex: number;
  sectionBreakEndIndex: number;
}) {
  // If the start index is undefined then we assume it's the first section break of the document which has a start index of 0
  const startIndex = sectionBreakStartIndex ?? 0;
  const endIndex = sectionBreakEndIndex;

  const sectionBreakStyle =
    sectionBreakElement.sectionStyle as docs_v1.Schema$SectionStyle;
  // Delete read only fields
  delete sectionBreakStyle.sectionType;
  const updateSectionBreakRequest = {
    updateSectionStyle: {
      sectionStyle: sectionBreakStyle,
      range: { startIndex, endIndex, segmentId: "" },
      fields: getStylingFieldMask(sectionBreakStyle),
    },
  };
  /**
   * If the start index does not exists then it must be the section break at the beginning of the document. In all other cases there should be a start index so we will insert a return a section break request with the appropriate styling.
   * Although, from what I've seen, there should only ever be one section break at the start of the document and one at the end (the one at the end is removed from the JSON parsing at the beginning of the script). This means there shouldn't really be a case where the there is a section break with a startIndex.
   */
  if (startIndex === 0) {
    // All documents must start with a section break insertion so that there can be an existing Paragraph object for other Structural Elements to be placed into. This is assuming there are no other section breaks in the document; I haven't found any places where a section break appears other than the first index of the document.
    return [
      {
        insertSectionBreak: {
          sectionType: "CONTINUOUS",
          // Because there are no existing pargraphs we need to insert it the sectionBreak at the end of the document body segment
          endOfSegmentLocation: {
            segmentId: "",
          },
        },
      },
      updateSectionBreakRequest,
    ];
  } else {
    return [
      {
        insertSectionBreak: {
          sectionType:
            // Assume that if the section type is unspecified then it is continuous
            sectionBreakElement.sectionStyle!.sectionType ===
            "SECTION_TYPE_UNSPECIFIED"
              ? "CONTINUOUS"
              : sectionBreakElement.sectionStyle!.sectionType,
          location: {
            segmentId: "",
            index: startIndex,
          },
        },
      },
      updateSectionBreakRequest,
    ];
  }
}

export { jsonToBatchUpdateRequest };
