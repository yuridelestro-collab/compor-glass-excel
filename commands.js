/* Compor Glass - comando direto na faixa de opções do Excel */
Office.onReady(() => {
  Office.actions.associate("clearColors", clearColors);
});

async function clearColors(event) {
  try {
    await Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.format.fill.clear();
      selectedRange.format.font.color = "#000000";
      await context.sync();
    });
  } catch (error) {
    console.error("Compor Glass - Limpar Cores:", error);
  } finally {
    event.completed();
  }
}
