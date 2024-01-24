//Pass the element id to the function

//    <button id="section" class="ms-Button">
//	      <span class="ms-Button-label" font-weight:bold>1.0</span>
//    </button>

$("#section").click(() => tryCatch(() => setStyle("Section Header")));
$("#subsection").click(() => tryCatch(() => setStyle("Section Subheading")));
$("#step1").click(() => tryCatch(() => setStyle("Step Level 1")));
$("#step2").click(() => tryCatch(() => setStyle("Step Level 2")));
$("#step3").click(() => tryCatch(() => setStyle("Step Level 3")));
// so on
// so forth


async function setStyle(name) {
  await Word.run(async (context) => {
    //Get selected range and expand it to include the whole first and last paragraphs
    var selection = context.document.getSelection().getRange();
    var firstParagraph = selection.paragraphs.getFirstOrNullObject();
    var lastParagraph = selection.paragraphs.getLastOrNullObject();
    var updatedSelection = selection.expandTo(firstParagraph.getRange()).expandTo(lastParagraph.getRange());
    //load the paragraphs and await sync
    updatedSelection.paragraphs.load();
    await context.sync();

    //console.log(updatedSelection.text)
    updatedSelection.style = name;
    await context.sync();

    //Move the cursor to the end of the selection
    updatedSelection.paragraphs
      .getLast()
      .getNextOrNullObject()
      .select("Start");
      return;
  });
}

/* Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
