import { useEffect, useState } from "react";
import { Button } from "./components/ui/button";
import { Label } from "./components/ui/label";
import { IpcRendererEvent } from "electron";
import { FaSpinner } from "react-icons/fa";

function App(): React.JSX.Element {
  const [message, setMessage] = useState<string>("");
  const [stopsReportfilePath, setstopsReportfilePath] = useState<string>("");
  const [workTimesReportfilPath, setWorkTimesReportfilPath] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);

  // Send request to main to open the file dialog
  const handleFileSelection = (fileType: string) => {
    console.log(fileType);
    window.electron.ipcRenderer.send("openFileDialog", fileType);
  };

  // Handle the response from the file selection
  useEffect(() => {
    const openFileDialogReply = (event: IpcRendererEvent, response): void => {
      if (response.success) {
        console.log("File selected successfully");
        if (response.fileType === "stopReport") {
          console.log("Stop report selected", response.results);
          setstopsReportfilePath(response.results);
        } else if (response.fileType === "workTimesReport") {
          console.log("Work Times report selected", response.results);
          setWorkTimesReportfilPath(response.results);
        }
      } else {
        console.log("An error has occurred");
        setMessage(response.error);
      }
    };

    // Add the listener on mount
    const openFileDialogReplyListener = window.electron.ipcRenderer.on(
      "openFileDialogReply",
      openFileDialogReply
    );

    // Remove the listener on unmount
    return (): void => {
      openFileDialogReplyListener();
    };
  }, []);

  // Send the request to main to combine the spradsheets
  const handleCombineSpreadsheets = (e: React.FormEvent<HTMLFormElement>): void => {
    e.preventDefault();
    console.log("Sending IPC request to main thread to combine spreadsheets");
    setIsLoading(true);
    setMessage("Combining spreadsheets. Please wait. .");
    if (stopsReportfilePath && workTimesReportfilPath) {
      window.electron.ipcRenderer.send(
        "combineSpreadsheets",
        stopsReportfilePath,
        workTimesReportfilPath
      );
    } else {
      setMessage("Please select stops and work times reports...");
      setIsLoading(false);
    }
  };

  // TODO => Handle the response after spreadsheets are combined
  useEffect(() => {
    const combineSpreadsheetsReply = (event: IpcRendererEvent, response): void => {
      if (response.success) {
        setMessage(response.message);
      } else {
        setMessage(response.error);
      }

      setIsLoading(false);
    };

    // Add the listener on mount
    const combineSpreadsheetsReplyListener = window.electron.ipcRenderer.on(
      "combineSpreadsheetsReply",
      combineSpreadsheetsReply
    );

    // Remove the listener on unmount
    return (): void => {
      combineSpreadsheetsReplyListener();
    };
  });

  return (
    <>
      <h1 className="text-4xl mb-4 m-6">InfoMatics Spreadsheet Combiner</h1>
      <form
        onSubmit={(e: React.FormEvent<HTMLFormElement>) => handleCombineSpreadsheets(e)}
        className="m-8 grid w-full max-w-sm items-center gap-6"
      >
        <Label htmlFor="stops-report-zip">Stops Report (.zip)</Label>
        <Button variant="outline" type="button" onClick={() => handleFileSelection("stopReport")}>
          Select .zip File
        </Button>
        {stopsReportfilePath && (
          <p>
            <b>Selected:</b>
            {stopsReportfilePath}
          </p>
        )}

        <Label htmlFor="work-time-zip">Work Times Report (.zip)</Label>
        <Button
          variant="outline"
          type="button"
          onClick={() => handleFileSelection("workTimesReport")}
        >
          Select .zip File
        </Button>
        {workTimesReportfilPath && (
          <p>
            <b>Selected:</b>
            {workTimesReportfilPath}
          </p>
        )}

        <div className="flex-row mt-3">
          <Button disabled={isLoading}>
            {isLoading ? (
              <span className="flex items-center gap-2">
                <FaSpinner className="animate-spin" />
                Combining...
              </span>
            ) : (
              "Combine Spreadsheets"
            )}
          </Button>
        </div>

        {message && (
          <div className="flex flex-row border p-6 border-black">
            <p>{message}</p>
          </div>
        )}
      </form>
    </>
  );
}

export default App;
