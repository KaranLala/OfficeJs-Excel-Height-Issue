/**
 * Gets the height and width of each cell in the used range of a worksheet.
 * @param {string} worksheetName The name of the worksheet.
 * @returns {Promise<{height: number, width: number}[][]>} A promise that resolves to the dimensions of each cell in the used range.
 */
export const getUsedRangeCellDimensions = async (): Promise<
  { height: number; width: number }[][]
> => {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load("rowIndex");
      range.load("rowCount");
      range.load("columnIndex");
      range.load("columnCount");
      await context.sync();

      let dimensions = [];
      for (var r = range.rowIndex; r < range.rowIndex + range.rowCount; r++) {
        let row = [];

        for (var c = range.columnIndex; c < range.columnIndex + range.columnCount; c++) {
          var cell = sheet.getCell(r, c);
          cell.load("height");
          cell.load("width");
          cell.load("address");
          row.push(cell);
        }

        dimensions.push(row);
      }

      await context.sync();

      return dimensions;
    });
  } catch (error) {
    console.error("Error: " + error);
    throw error;
  }
};

export default interface MergedArea {
  col: number;
  row: number;
  colCount: number;
  rowCount: number;
  width: number;
  height: number;
  values: string[][];
}

/**
 * Gets all merged areas within the used range of a specified worksheet.
 * @param {string} worksheetName The name of the worksheet.
 * @returns {Promise<string[]>} A promise that resolves to an array of address strings for merged areas.
 */
export const getMergedAreas = async (): Promise<MergedArea[]> => {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load("columnIndex");
      range.load("rowIndex");

      const mergedAreas = range.getMergedAreasOrNullObject();
      mergedAreas.load("address");
      mergedAreas.load("areas");
      mergedAreas.load("cellCount");
      mergedAreas.load("areaCount");
      await context.sync();

      if (mergedAreas.isNullObject) {
        return [];
      } else {
        console.log(mergedAreas.areas.items);
        return mergedAreas.areas.items.map((a: any) => ({
          col: a.columnIndex - range.columnIndex,
          row: a.rowIndex - range.rowIndex,
          colCount: a.columnCount,
          rowCount: a.rowCount,
          height: a.height,
          width: a.width,
          values: a.values,
        }));
      }
    });
  } catch (error) {
    console.error("Error: " + error);
    throw error;
  }
};
