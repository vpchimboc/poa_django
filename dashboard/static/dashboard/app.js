const POA = JSON.parse(document.getElementById("poa-data").textContent);
const CATS = ["Unidad Administrativa", "Unidad Académica", "Carrera"];

const pct = (v) => (v === null || v === undefined ? "—" : Math.round(v * 100) + "%");
const nivel = (v) => ((v ?? 0) < 0.35 ? "critico" : (v ?? 0) < 0.7 ? "medio" : "alto");
const nivelLabel = { critico: "Crítico", medio: "En proceso", alto: "Satisfactorio" };
const color = { critico: "#c0392b", medio: "#d98324", alto: "#0f8a8a" };
const esc = (s) =>
  String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

let estado = { q: "", cat: "Todas", selected: POA.unidades[0]?.id ?? null };

function bar(v, h) {
  const n = nivel(v);
  return `<div class="bar" style="height:${h || 6}px"><i style="width:${Math.min(100, Math.round((v ?? 0) * 100))}%;background:${color[n]}"></i></div>`;
}
function badge(v) {
  const n = nivel(v);
  return `<span class="badge lvl-${n}">${nivelLabel[n]} · ${pct(v)}</span>`;
}

function filtradas() {
  const term = estado.q.trim().toLowerCase();
  return POA.unidades.filter((u) => {
    const okCat = estado.cat === "Todas" || u.categoria === estado.cat;
    const hay = [u.nombre, u.titulo, u.responsable, u.categoria].filter(Boolean).join(" ").toLowerCase();
    return okCat && (!term || hay.includes(term));
  });
}

function resumen(unidades) {
  const acts = unidades.flatMap((u) => u.actividades);
  const con = acts.filter((a) => a.avance !== null && a.avance !== undefined);
  return {
    unidades: unidades.length,
    actividades: acts.length,
    completadas: con.filter((a) => a.avance >= 1).length,
    sinIniciar: con.filter((a) => a.avance === 0).length,
    observaciones: acts.filter((a) => a.observacion).length,
    promedio: unidades.length ? unidades.reduce((s, u) => s + u.avance, 0) / unidades.length : 0,
  };
}

function renderDetalle(u) {
  const el = document.getElementById("detalle");
  if (!u) {
    el.innerHTML = '<p class="empty">Selecciona una unidad para ver su detalle.</p>';
    return;
  }
  const grupos = new Map();
  u.actividades.forEach((a) => {
    const k = a.objetivo || "Actividades planificadas";
    if (!grupos.has(k)) grupos.set(k, []);
    grupos.get(k).push(a);
  });

  el.innerHTML = `
    <div class="row">
      <div>
        <h2>${esc(u.nombre)}</h2>
        <p class="muted">${esc(u.titulo || u.categoria)}${u.responsable ? " · " + esc(u.responsable) : ""}</p>
      </div>
      ${badge(u.avance)}
    </div>
    ${bar(u.avance, 8)}
    <p class="muted" style="margin-top:10px">${u.totalActividades} actividades registradas</p>
    ${[...grupos.entries()]
      .map(
        ([obj, acts]) => `<div class="obj"><h3>${esc(obj)}</h3>${acts
          .map(
            (a) => `<div class="act">
              <div class="row"><span class="tarea">${esc(a.tarea)}</span>${badge(a.avance)}</div>
              <div class="meta">
                ${a.responsable ? `<span><b>Responsable:</b> ${esc(a.responsable)}</span>` : ""}
                ${a.fecha ? `<span><b>Entrega:</b> ${esc(a.fecha)}</span>` : ""}
                ${a.prioridad ? `<span><b>Prioridad:</b> ${esc(a.prioridad)}</span>` : ""}
                ${a.entregable ? `<span><b>Entregable:</b> ${esc(a.entregable)}</span>` : ""}
              </div>
              ${bar(a.avance)}
              ${a.observacion ? `<div class="obs"><strong>Observación:</strong> ${esc(a.observacion)}</div>` : ""}
            </div>`,
          )
          .join("")}</div>`,
      )
      .join("")}`;
}

function render() {
  const list = filtradas();
  const base = list.length ? list : POA.unidades;
  const r = resumen(base);

  document.getElementById("kpi-promedio").textContent = pct(r.promedio);
  document.getElementById("kpi-unidades").textContent = `${r.unidades} unidades en el filtro actual`;
  document.getElementById("kpi-actividades").textContent = r.actividades;
  document.getElementById("kpi-completadas").textContent = `${r.completadas} al 100% de ejecución`;
  document.getElementById("kpi-sin").textContent = r.sinIniciar;
  document.getElementById("kpi-obs").textContent = r.observaciones;

  document.getElementById("count").textContent = list.length;
  document.getElementById("unidades").innerHTML = list.length
    ? list
        .map(
          (u) => `<li><button data-id="${esc(u.id)}" class="${u.id === estado.selected ? "active" : ""}">
            <div class="row"><span class="name">${esc(u.nombre)}</span><strong>${pct(u.avance)}</strong></div>
            <p class="muted" style="margin:2px 0 0">${esc(u.categoria)}</p>
            ${bar(u.avance)}
          </button></li>`,
        )
        .join("")
    : '<li class="empty">Sin unidades para esta búsqueda.</li>';

  document.querySelectorAll("#unidades button").forEach((b) =>
    b.addEventListener("click", () => {
      estado.selected = b.dataset.id;
      render();
    }),
  );

  const sel = POA.unidades.find((u) => u.id === estado.selected) || list[0];
  renderDetalle(sel);

  document.getElementById("chart").innerHTML = [...base]
    .sort((a, b) => b.avance - a.avance)
    .map(
      (u) => `<div class="cbar"><span title="${esc(u.nombre)}">${esc(u.nombre)}</span>
        <span class="track"><i style="width:${Math.min(100, Math.round(u.avance * 100))}%;background:${color[nivel(u.avance)]}"></i></span>
        <strong>${pct(u.avance)}</strong></div>`,
    )
    .join("");
}

document.getElementById("proyectos").innerHTML = POA.proyectos
  .map(
    (p) => `<li>
      <div class="row"><h3>${esc(p.nombre)}</h3>${badge(p.avance)}</div>
      ${p.descripcion ? `<p class="muted">${esc(p.descripcion)}</p>` : ""}
      ${bar(p.avance)}
      ${p.responsable ? `<p class="muted" style="margin:8px 0 0">${esc(p.responsable)}</p>` : ""}
      ${p.observacion ? `<div class="obs">${esc(p.observacion)}</div>` : ""}
    </li>`,
  )
  .join("");

document.getElementById("q").addEventListener("input", (e) => {
  estado.q = e.target.value;
  render();
});
document.querySelectorAll(".chip").forEach((c) =>
  c.addEventListener("click", () => {
    document.querySelectorAll(".chip").forEach((x) => x.classList.remove("active"));
    c.classList.add("active");
    estado.cat = c.dataset.cat;
    render();
  }),
);

render();
