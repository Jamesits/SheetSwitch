import * as React from "react";
import { useState, useEffect } from "react";
import usePromise from 'react-use-promise';
import { useExcelEvent } from "./ExcelKit";
import classNames from 'classnames/bind';

const styles = require("./SheetList.module.css");
let cx = classNames.bind(styles);

const loadSheets = async () => Excel.run(async function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync();
    return sheets.items.map((item) => item.name);
})

const getActiveSheet = async () => Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();
    return sheet.name;
})

const setActiveSheet = (i: string, callback: (string) => void) => {
    console.log("Switching to", i);
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem(i);
        sheet.activate();
        sheet.load("name");

        return context.sync()
            .then(function () {
                console.log(`The active worksheet is "${sheet.name}"`);
                callback(sheet.name);
            });
    }).catch(() => { console.log("error") });
}

const SheetList: React.FC = () => {
    const [activeSheetName, setActiveSheetName] = useState("");
    const [worksheets, setWorksheets] = useState([]);
    const [result, error, state] = usePromise(async () => setWorksheets(await loadSheets()), []);
    usePromise(async () => setActiveSheetName(await getActiveSheet()), []);

    useExcelEvent(
        (context) => context.workbook.worksheets.onAdded,
        async () => setWorksheets(await loadSheets()),
        []
    );

    useExcelEvent(
        (context) => context.workbook.worksheets.onChanged,
        async () => setWorksheets(await loadSheets()),
        []
    );

    useExcelEvent(
        (context) => context.workbook.worksheets.onDeleted,
        async () => setWorksheets(await loadSheets()),
        []
    );

    useExcelEvent(
        (context) => context.workbook.worksheets.onActivated,
        async () => setActiveSheetName(await getActiveSheet()),
        []
    );

    useEffect(() => {
        console.log("useEffect", worksheets);
    });

    if (error) return <div>Error</div>
    if (state !== "resolved") return <div>Loading</div>

    return (
        <div>
            <ol className={styles["sheet-list"]}>
                {worksheets.map((value, key) => 
                    <li key={key} onClick={setActiveSheet.bind(undefined, value, setActiveSheetName)}
                        className={cx('sheet-list-item', {
                            'sheet-list-item-active':  activeSheetName === value,
                        })}
                    >{value}</li>
                )}
            </ol>
        </div>
    )
}

export default SheetList;
