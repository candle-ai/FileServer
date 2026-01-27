import "./modulepreload-polyfill.DaKOjhqt.js";
Office.onReady(() => {
  console.log("Outlook add-in commands loaded");
});
function openTaskpane() {
  Office.ribbon.requestUpdate({
    tabs: [{
      id: "TabDefault",
      groups: [{
        id: "CandleGroup",
        controls: [{
          id: "TaskpaneButton",
          enabled: true
        }]
      }]
    }]
  });
}
function analyzeEmail() {
  console.log("Analyze email command triggered");
}
function generateDraft() {
  console.log("Generate draft command triggered");
}
window.openTaskpane = openTaskpane;
window.analyzeEmail = analyzeEmail;
window.generateDraft = generateDraft;
