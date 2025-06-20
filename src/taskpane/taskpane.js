import {wordList} from "./data.js";
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = enableSuggestionListener;
    document.getElementById("stop").onclick = disableSuggestionListener;
    console.log("Office is ready.");
  }
});
// Enable the listener
async function enableSuggestionListener() {
  await Word.run(async (context) => {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChange,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Suggestion listener attached.");
        } else {
          console.error("Failed to attach listener:", result.error.message);
        }
      }
    );
  });
}

// Disable the listener
async function disableSuggestionListener() {
  await Word.run(async (context) => {
    Office.context.document.removeHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      { handler: onSelectionChange },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Suggestion listener removed.");
          document.getElementById("suggestions").style.display = "none";
        } else {
          console.error("Failed to remove listener:", result.error.message);
        }
      }
    );
  });
}

// The actual handler function
async function onSelectionChange() {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const text = selection.text.trim();
    if (!text) {
      document.getElementById("suggestions").style.display = "none";
      return;
    }
    const matches = searchDataset(text);
    renderDropdown(matches, text);
  });
}

async function insertSuggestion(suggestion) {
  return Word.run(async (context) => {
    const range = context.document.getSelection();
    range.load("text");
    await context.sync();
    
    const currentText = range.text;
    console.log("Current text:", currentText);
    const updatedText = currentText.replace(new RegExp(`${range.text}$`), suggestion);
    console.log("Updated text:", updatedText);
    range.insertText(updatedText, Word.InsertLocation.replace);
    await context.sync();

    document.getElementById("suggestions").style.display = "none";
  });
}

function searchDataset(prefix) {
  if (!prefix) return [];
  //get the size of wordList
  const size = wordList.length;
  prefix=prefix.toLowerCase();
  //loop through the wordList and find matches
  const matches = [];
  for (let i = 0; i < size; i++) {
    const word = wordList[i].toLowerCase();
    if (word.startsWith(prefix)) {
      matches.push(word);
    }
  }
  return matches;
}

function renderDropdown(matches) {
  const container = document.getElementById("suggestions");
  if (!container) {
    console.warn("Suggestions container not found.");
    return;
  }

  container.innerHTML = "";

  if (!matches.length) {
    container.style.display = "none";
    return;
  }

  matches.forEach((match) => {
    const item = document.createElement("div");
    item.className = "suggestion";
    item.textContent = match;
    item.onclick = () => insertSuggestion(match);
    container.appendChild(item);
  });

  container.style.display = "block";
}
