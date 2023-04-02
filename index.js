/**
 * entry
 */
async function main() {
  logseq.App.registerCommandPalette({
    key: "todos-export-excel",
    label: "Export TODOs to Excel"
  }, async () => {
    const tasks = await logseq.App.q('(task TODO)');

    const data = tasks.map(task => {
      // Use the uuid, marker, content, id
      return {
        uuid: task.uuid,
        marker: task.marker,
        content: task.content,
        id: task.id
      }
    }
    );
    // Use XLSX to export data
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);

    // Auto size uuid column
    ws['!cols'] = [
      { wch: 36 },
      { wch: 5 },
      { wch: 100 },
      { wch: 5 }
    ];

    const heights = [];
    for (let i = 0; i < data.length + 1; i++) {
      heights.push({ hpx: 20 });
    }

    ws['!rows'] = heights;

    XLSX.utils.book_append_sheet(wb, ws, "TODOs");
    XLSX.writeFile(wb, "todos.xlsx");
  });
}

// bootstrap
logseq.ready(main).catch(console.error)

