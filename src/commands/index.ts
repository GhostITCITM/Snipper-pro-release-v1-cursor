import { modeManager } from "../helpers/mode";

Office.onReady(() => {
  console.log("Commands module loaded");
});

async function textSnip(): Promise<void> {
  try {
    await modeManager.setMode("text");
    console.log("Text snip mode activated");
  } catch (error) {
    console.error("Error activating text snip:", error);
  }
}

async function sumSnip(): Promise<void> {
  try {
    await modeManager.setMode("sum");
    console.log("Sum snip mode activated");
  } catch (error) {
    console.error("Error activating sum snip:", error);
  }
}

async function tableSnip(): Promise<void> {
  try {
    await modeManager.setMode("table");
    console.log("Table snip mode activated");
  } catch (error) {
    console.error("Error activating table snip:", error);
  }
}

async function validationSnip(): Promise<void> {
  try {
    await modeManager.setMode("validation");
    console.log("Validation snip mode activated");
  } catch (error) {
    console.error("Error activating validation snip:", error);
  }
}

async function exceptionSnip(): Promise<void> {
  try {
    await modeManager.setMode("exception");
    console.log("Exception snip mode activated");
  } catch (error) {
    console.error("Error activating exception snip:", error);
  }
}

Office.actions.associate("textSnip", textSnip);
Office.actions.associate("sumSnip", sumSnip);
Office.actions.associate("tableSnip", tableSnip);
Office.actions.associate("validationSnip", validationSnip);
Office.actions.associate("exceptionSnip", exceptionSnip);
