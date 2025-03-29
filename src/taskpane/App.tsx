import * as React from "react";

import { useState, useEffect } from "react";
import { Document, Page, View, Text, StyleSheet, PDFViewer } from "@react-pdf/renderer";
import { getMergedAreas, getUsedRangeCellDimensions } from "./office";

import "./styles.css";

const styles = StyleSheet.create({
  page: {
    flexDirection: "row",
    backgroundColor: "#E4E4E4",
  },
  section: {
    margin: 10,
    padding: 10,
    flexGrow: 1,
  },
});

const App = () => {
  const [usedCellDimensions, setUsedCellDimensions] = useState<any>();
  const [mergedAreas, setMergedAreas] = useState<any>();

  const refresh = async () => {
    setUsedCellDimensions(await getUsedRangeCellDimensions());
    setMergedAreas(await getMergedAreas());
  };

  useEffect(() => {
    refresh();
  }, []);

  console.log(usedCellDimensions);
  console.log(mergedAreas);

  return (
    <div className="bg-gray-100 h-full">
      <div className="flex items-center justify-center gap-2 mt-2">
        <button
          className="border border-gray-300 rounded-md p-2 bg-blue-500 hover:bg-blue-600 text-white shadow-sm"
          onClick={refresh}
        >
          Refresh Data for Active Sheet
        </button>
      </div>
      <div className="grid grid-cols-2 gap-5 mt-2">
        <div>
          <div className="text-lg font-bold">Data by Cell Dimensions</div>
          <div>
            {usedCellDimensions?.map((row: any) => {
              let cell = row[0];
              return (
                <div key={cell.address}>
                  <div>
                    {cell.address}: height = {cell.height}
                  </div>
                </div>
              );
            })}
          </div>
          <div>
            <span className="font-bold">Total: </span>
            <span>
              {usedCellDimensions?.reduce((acc: number, row: any) => {
                return acc + row[0].height;
              }, 0)}
            </span>
          </div>
        </div>
        <div>
          <div className="text-lg font-bold">Data by Merged Areas</div>
          <div>
            <span className="font-bold">Total: </span>
            <span>{mergedAreas?.[0].height}</span>
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;
