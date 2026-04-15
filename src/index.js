function setStatus(message, isError) {
  const status = document.getElementById("hostStatus");
  if (!status) {
    return;
  }
  status.textContent = message;
  status.style.background = isError ? "#f7d8d2" : "#edf4f4";
  status.style.color = isError ? "#6a1b16" : "#134243";
}

function init() {
  const main = document.getElementById("mainDiv");
  if (!main) {
    return;
  }

  const intro = document.createElement("p");
  intro.className = "hello-intro";
  intro.textContent =
    "This add-in is running. Use the button below to insert a paragraph into the document.";

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "primary";
  btn.textContent = "Insert “Hello World” into the document";

  btn.addEventListener("click", () => {
    if (typeof Word === "undefined") {
      setStatus("Word API is not available here. Open this add-in in Word.", true);
      return;
    }
    Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph("Hello World", Word.InsertLocation.end);
      await context.sync();
      setStatus("Inserted “Hello World” at the end of the document.", false);
    }).catch((error) => {
      setStatus(error.message || "Failed to insert text", true);
    });
  });

  main.appendChild(intro);
  main.appendChild(btn);
}

Office.onReady((info) => {
  const inWord = info.host === Office.HostType.Word;
  setStatus(
    inWord
      ? "Connected to Word"
      : "Ready — open this add-in in Microsoft Word to use Insert.",
    false
  );
  init();
});
