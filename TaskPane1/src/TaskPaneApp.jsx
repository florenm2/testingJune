import React, { useEffect } from 'react';
import './App.css';
import * as Office from '@microsoft/office-js';

function TaskPaneApp() {
  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Word || info.host === Office.HostType.PowerPoint) {
        document.getElementById("insertIframeBtn").onclick = insertIframe;
      }
    });
  }, []);

  const insertIframe = () => {
    const iframeUrl = "https://path/to/your/iframe.html"; // Update this to the URL of your HTML page

    Office.context.document.body.insertHtml(
      `<iframe src="\${iframeUrl}" width="600" height="400"></iframe>`,
      { coercionType: Office.CoercionType.Html },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Iframe inserted successfully!");
        } else {
          console.error("Error inserting iframe: " + asyncResult.error.message);
        }
      }
    );
  };

  return (
    <div className="App">
      <h1>Hello, world!</h1>
      <button id="insertIframeBtn">Insert iFrame</button>
    </div>
  );
}

export default TaskPaneApp;
