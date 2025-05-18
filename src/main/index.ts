import { app, shell, BrowserWindow, ipcMain, IpcMainEvent, dialog } from "electron";
import path, { join } from "path";
import { electronApp, optimizer, is } from "@electron-toolkit/utils";
import icon from "../../resources/icon.png?asset";

import fs from "fs";
import unzipper from "unzipper";
import xlsx from "xlsx";

type RowObject = Record<string, any>; // Each row as an object with string keys

function createWindow(): void {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 900,
    height: 670,
    show: false,
    autoHideMenuBar: true,
    ...(process.platform === "linux" ? { icon } : {}),
    webPreferences: {
      preload: join(__dirname, "../preload/index.js"),
      sandbox: false
    }
  });

  mainWindow.on("ready-to-show", () => {
    mainWindow.show();
  });

  mainWindow.webContents.setWindowOpenHandler((details) => {
    shell.openExternal(details.url);
    return { action: "deny" };
  });

  // HMR for renderer base on electron-vite cli.
  // Load the remote URL for development or the local html file for production.
  if (is.dev && process.env["ELECTRON_RENDERER_URL"]) {
    mainWindow.loadURL(process.env["ELECTRON_RENDERER_URL"]);
  } else {
    mainWindow.loadFile(join(__dirname, "../renderer/index.html"));
  }
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(() => {
  // Set app user model id for windows
  electronApp.setAppUserModelId("com.electron");

  // Default open or close DevTools by F12 in development
  // and ignore CommandOrControl + R in production.
  // see https://github.com/alex8088/electron-toolkit/tree/master/packages/utils
  app.on("browser-window-created", (_, window) => {
    optimizer.watchWindowShortcuts(window);
  });

  const handleOpenFileDialog = async (event: IpcMainEvent, fileType: string) => {
    try {
      const result = await dialog.showOpenDialog({
        properties: ["openFile"],
        filters: [
          { name: "Zip Files", extensions: ["*.zip"] },
          { name: "All Files", extensions: ["*"] }
        ]
      });

      if (result.canceled) {
        event.reply("openFileDialogReply", {
          success: false,
          error: "File selection cancelled"
        });
        return;
      }

      // Return the file path to the render process
      console.log("[+] File path: ", result.filePaths[0]);
      event.reply("openFileDialogReply", {
        success: true,
        results: result.filePaths[0],
        fileType: fileType
      });
    } catch (error) {
      console.log("[-] An error occurred whle trying to select a file: ", error);
      event.reply("openFileDialogReply", { success: false, error: error });
    }
  };

  const unzipFile = async (inputFile: string, outputPath: string): Promise<boolean> => {
    try {
      await fs.promises.mkdir(outputPath, { recursive: true }); // Ensure output dir exists

      return new Promise((resolve, reject) => {
        const readStream = fs.createReadStream(inputFile);
        const extractStream = unzipper.Extract({ path: outputPath });

        readStream.pipe(extractStream);

        extractStream.on("close", () => {
          console.log("Unzipped successfully to", outputPath);
          resolve(true);
        });

        extractStream.on("error", (err) => {
          console.error("Unzip failed:", err);
          reject(false);
        });
      });
    } catch (error) {
      console.log("An error occurred:", error);
      return false;
    }
  };

  // Deletes the directory passed in
  const deleteDir = (dirPath: string): Promise<void> => {
    return new Promise((resolve, reject) => {
      fs.rm(dirPath, { recursive: true, force: true }, (err) => {
        if (err) reject(err);
        else resolve();
      });
    });
  };

  // Reads the directory and returns an array of strings
  const readDir = (dirPath: string): Promise<string[]> => {
    return new Promise((resolve, reject) => {
      fs.readdir(dirPath, (err, files) => {
        if (err) {
          reject(err);
        } else {
          resolve(files);
        }
      });
    });
  };

  const excelFractionToHMS = (fraction: number): string => {
    const totalSeconds = Math.round(fraction * 24 * 3600);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;

    // Pad with leading zeros if needed
    const hh = hours.toString().padStart(2, "0");
    const mm = minutes.toString().padStart(2, "0");
    const ss = seconds.toString().padStart(2, "0");

    return `${hh}:${mm}:${ss}`;
  };

  const writeCombinedDataToFile = (combinedData: Record<string, any>[], outputFilePath: string) => {
    // Convert JSON array to a worksheet
    const worksheet = xlsx.utils.json_to_sheet(combinedData);

    // Create a new workbook and append the worksheet
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, "Combined Report");

    // Write the workbook to a file
    xlsx.writeFile(workbook, outputFilePath);

    console.log(`Combined report written to: ${outputFilePath}`);
  };

  const combineStopReport = (stopReportFiles: string[], stopsReportExtractPath) => {
    let combinedData: RowObject[] = [];
    let isFirstFile = true;

    for (const fileName of stopReportFiles) {
      const fullPath = path.join(stopsReportExtractPath, fileName);
      const workbook = xlsx.readFile(fullPath);
      const sheetNames = workbook.SheetNames;
      const firstSheet = workbook.Sheets[sheetNames[0]];

      const options = {
        range: 5, // skip first 5 rows
        defval: ""
      };

      if (isFirstFile) {
        const data = xlsx.utils.sheet_to_json<RowObject>(firstSheet, options);
        combinedData = combinedData.concat(data);
        isFirstFile = false;
      } else {
        const rawRows = xlsx.utils.sheet_to_json(firstSheet, { header: 1, range: 5 }) as any[][];
        const headers = Object.keys(combinedData[0]);
        const dataRows = rawRows.slice(1);
        const dataObjects = dataRows.map((row) =>
          headers.reduce(
            (obj, key, index) => {
              obj[key] = row[index];
              return obj;
            },
            {} as Record<string, any>
          )
        );
        combinedData = combinedData.concat(dataObjects);
      }
    }

    // Adjust the times from excel time fraction to HH:MM:SS representation
    for (const row of combinedData) {
      row["Parking time"] = excelFractionToHMS(row["Parking time"]);
      row["Ignition on"] = excelFractionToHMS(row["Ignition on"]);
      row["Engine on"] = excelFractionToHMS(row["Engine on"]);
    }

    // Set the output path (using the date to avoid accidentally overriding)
    const currentDateTime = new Date().toISOString().replace(/[:.]/g, "-");
    const outputPath = path.join(
      app.getPath("downloads"),
      `combinedStopsReport-${currentDateTime}.xlsx`
    );

    writeCombinedDataToFile(combinedData, outputPath);
  };

  const renameWorkTimeReportHeaders = (data: RowObject[]): RowObject[] => {
    const headerMap: Record<string, string> = {
      "": "Date",
      _1: "Licence plate",
      _2: "Vehicle",
      _3: "Vehicle Title",
      _4: "Total Driving Time"
      // _5: "Distance"
    };

    return data.map((row) => {
      const renamedRow: RowObject = {};
      for (const key in row) {
        const newKey = headerMap[key] ?? key; // use mapped key or original
        renamedRow[newKey] = row[key];
      }
      return renamedRow;
    });
  };

  const excelMinutesToHMS = (minutes: number): string => {
    const totalSeconds = Math.round(minutes * 60);
    const hours = Math.floor(totalSeconds / 3600);
    const minutesPart = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return `${String(hours).padStart(2, "0")}:${String(minutesPart).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  };

  const combineWorkTimesReport = (workTimeReportFiles: string[], workTimesExtractPath) => {
    let combinedData: RowObject[] = [];
    let isFirstFile = true;

    for (const fileName of workTimeReportFiles) {
      const fullPath = path.join(workTimesExtractPath, fileName);
      const workbook = xlsx.readFile(fullPath);
      const sheetNames = workbook.SheetNames;
      const firstSheet = workbook.Sheets[sheetNames[0]];

      const options = {
        range: 4, // Skip first 4 rows
        defval: ""
      };

      if (isFirstFile) {
        let data = xlsx.utils.sheet_to_json<RowObject>(firstSheet, options);
        data = renameWorkTimeReportHeaders(data);
        combinedData = combinedData.concat(data);
        isFirstFile = false;
      } else {
        const rawRows = xlsx.utils.sheet_to_json(firstSheet, { header: 1, range: 4 }) as any[][];
        const headers = Object.keys(combinedData[0]);
        const dataRows = rawRows.slice(1);
        const dataObjects = dataRows.map((row) =>
          headers.reduce(
            (obj, key, index) => {
              obj[key] = row[index];
              return obj;
            },
            {} as Record<string, any>
          )
        );
        combinedData = combinedData.concat(dataObjects);
      }
    }

    // Remove the unecessary data
    combinedData = combinedData.map((row) => {
      const {
        "Driving time_1": _,
        "Total Driving Time": __,
        _5: ___,
        Distance_1: ____,
        ...rest
      } = row;
      return rest;
    });

    // Fix the formatting of excessive idling
    for (const row of combinedData) {
      if (isNaN(row["Excessive idling"])) {
        row["Excessive idling"] = "00:00:00";
      } else {
        row["Excessive idling"] = excelMinutesToHMS(row["Excessive idling"]);
      }
    }

    // Set the output path (using the date to avoid accidentally overriding)
    const currentDateTime = new Date().toISOString().replace(/[:.]/g, "-");
    const outputPath = path.join(
      app.getPath("downloads"),
      `combinedWorkTimesReport-${currentDateTime}.xlsx`
    );

    writeCombinedDataToFile(combinedData, outputPath);
  };

  const handleCombineSpreadsheets = async (
    event: IpcMainEvent,
    stopsReportfilePath: string,
    workTimesReportfilePath: string
  ) => {
    // Set the paths the the user data directory. This will get overwritten everytime the app runs and so should not cause an issue
    const stopsReportExtractPath = path.join(app.getPath("userData"), "stopsReport");
    const workTimesExtractPath = path.join(app.getPath("userData"), "workTimesReport");

    // Delete the files already in the extract directory
    await deleteDir(stopsReportExtractPath);
    await deleteDir(workTimesExtractPath);

    // Unzip the files into the extract directory
    const stopReportExtractSuccess = await unzipFile(stopsReportfilePath, stopsReportExtractPath);
    const workTimesReportExtractSuccess = await unzipFile(
      workTimesReportfilePath,
      workTimesExtractPath
    );

    // Return an error if we can't unzip for some reason
    if (!stopReportExtractSuccess || !workTimesReportExtractSuccess) {
      event.reply("combineSpreadsheetsReply", {
        success: false,
        error: "Issue unzipping the files"
      });
    }

    // Combine stop reports
    const stopReportFileNames = await readDir(stopsReportExtractPath);
    combineStopReport(stopReportFileNames, stopsReportExtractPath);

    // Combine Work times reports
    const workTimeReportFileNames = await readDir(workTimesExtractPath);
    combineWorkTimesReport(workTimeReportFileNames, workTimesExtractPath);

    event.reply("combineSpreadsheetsReply", {
      success: true,
      message: "Spreadsheets combined. Please check output directory"
    });
  };

  // Receive IPC from render process to open file dialog
  ipcMain.on("openFileDialog", (event: IpcMainEvent, fileType: string) =>
    handleOpenFileDialog(event, fileType)
  );

  // Recive IPC from render process to combine the spreadsheets
  ipcMain.on(
    "combineSpreadsheets",
    (event: IpcMainEvent, stopsReportfilePath: string, workTimesReportfilePath) =>
      handleCombineSpreadsheets(event, stopsReportfilePath, workTimesReportfilePath)
  );

  createWindow();

  app.on("activate", function () {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.
