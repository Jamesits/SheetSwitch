import * as React from "react";
import { useState, useEffect } from "react";
import usePromise from 'react-use-promise';
import { useExcelEvent } from "./ExcelKit";

// const loadExample = function <T>(batch: (context: Excel.RequestContext) => Promise<T>) {
//     return new Promise((resolve, reject) => {
//         Excel.run((context) => {
//             return batch(context).then(resolve, reject)
//         }).then(resolve, reject)
//     });
// }

const loadSheets = async () => Excel.run(async function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();
    return sheets.items.map((item) => item.name);
})

const setActiveSheet = (i: string) => {
    console.log("Switching to", i);
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem(i);
        sheet.activate();
        sheet.load("name");

        return context.sync()
            .then(function () {
                console.log(`The active worksheet is "${sheet.name}"`);
            });
    }).catch(() => { console.log("error") });
}

const SheetList: React.FC = () => {
    const [worksheets, setWorksheets] = useState([]);
    const [result, error, state] = usePromise(async () => setWorksheets(await loadSheets()), []);

    useExcelEvent(
        (context) => context.workbook.worksheets.onAdded,
        async () => {
            const sheets = await loadSheets();
            console.log(sheets);
            setWorksheets(sheets);
        }, 
        []
    );

    useEffect(() => {
        console.log("useEffect", worksheets);
    });

    if (error) return <div>Error</div>
    if (state !== "resolved") return <div>Loading</div>

    return (
        <div>
            Sheets: 
            {worksheets.length}
            <ol>
                {worksheets.map((value, key) => <li key={key} onClick={setActiveSheet.bind(undefined, value)}>{value}</li>)}
            </ol>
        </div>
    )
}

export default SheetList;
