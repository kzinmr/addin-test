/* global document, Office, module, require */

const run = async () => {
  try{
    await Word.run(async (context) => {
      context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
      await context.sync();
    });
  } catch(error) {
    console.log(error);
  };
}

Office.onReady((info: Office.InitializationContext) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = run;
  }
});