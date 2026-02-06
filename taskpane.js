let data = [
  { Objekt: "BVH Musterstraße", ERGO: "1001", CUP: "A1", CIG: "X01" },
  { Objekt: "BVH Hauptplatz", ERGO: "1002", CUP: "A2", CIG: "X02" },
  { Objekt: "BVH Gewerbepark", ERGO: "1003", CUP: "A3", CIG: "X03" }
];

Office.onReady(() => {
  render(data);
  setupSearch();
});

function setupSearch() {
  document.getElementById("search").oninput = e => {
    const t = e.target.value.toLowerCase();
    render(data.filter(d => d.Objekt.toLowerCase().includes(t)));
  };
}

function render(items) {
  const list = document.getElementById("list");

  while (list.firstChild) {
    list.removeChild(list.firstChild);
  }

  items.forEach(i => {
    const o = document.createElement("option");
    o.value = JSON.stringify(i);
    o.textContent = `${i.Objekt} – ERGO ${i.ERGO}`;
    list.appendChild(o);
  });
}

document.getElementById("insert").onclick = () => {
  const val = document.getElementById("list").value;
  if (!val) return;

  const b = JSON.parse(val);

  Office.context.mailbox.item.body.setSelectedDataAsync(
`Objekt: ${b.Objekt}
ERGO: ${b.ERGO}
CUP: ${b.CUP}
CIG: ${b.CIG}`,
    { coercionType: Office.CoercionType.Text }
  );
};
