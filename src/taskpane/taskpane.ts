import "./taskpane.css";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    initialize();
  }
});

function initialize(): void {
  console.log("PowerPoint Toolbox initialized");
}
