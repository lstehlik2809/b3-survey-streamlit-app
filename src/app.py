from __future__ import annotations

import base64
import html
import json
import tempfile
import traceback
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

try:
    from b3_analysis import B3InputError, generate_b3_report
except ModuleNotFoundError:
    from src.b3_analysis import B3InputError, generate_b3_report


st.set_page_config(page_title="B3 report", layout="wide")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_INPUT_PATH = PROJECT_ROOT / "B3Inputs.xlsx"
REPORT_STATE_KEYS = [
    "report_bytes",
    "preview_html",
    "positive_network_html",
    "negative_network_html",
    "layout_engine",
    "stats",
]


def main() -> None:
    if "upload_key" not in st.session_state:
        st.session_state.upload_key = 0

    st.title("B3 report")
    st.write(
        "Nahrajte vyplněný Excelovský soubor v požadovaném formátu "
        "(viz ukázkový soubor), vygenerujte a stáhněte report, případně "
        "prozkoumejte interaktivní grafické výstupy."
    )

    if SAMPLE_INPUT_PATH.exists():
        st.download_button(
            "Stáhnout ukázkový soubor",
            data=SAMPLE_INPUT_PATH.read_bytes(),
            file_name="B3Inputs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    uploaded_file = st.file_uploader(
        "Nahrát soubor",
        type=["xlsx", "xlsm"],
        key=f"uploaded_file_{st.session_state.upload_key}",
    )
    generate_clicked = st.button(
        "Vygenerovat report",
        type="primary",
        disabled=uploaded_file is None,
    )

    if generate_clicked and uploaded_file is not None:
        _generate_report(uploaded_file)

    if "report_bytes" in st.session_state:
        _render_result()


def _generate_report(uploaded_file) -> None:
    with st.spinner("Generuji report..."):
        try:
            with tempfile.TemporaryDirectory(prefix="b3-report-") as temp_dir:
                temp_path = Path(temp_dir)
                suffix = Path(uploaded_file.name).suffix.lower()
                workbook_path = temp_path / f"uploaded_workbook{suffix}"
                workbook_path.write_bytes(uploaded_file.getvalue())

                result = generate_b3_report(workbook_path, temp_path)
                report_bytes = result.report_path.read_bytes()

                st.session_state.report_bytes = report_bytes
                st.session_state.preview_html = _build_preview_html(result)
                st.session_state.positive_network_html = _build_interactive_network_html(
                    result.positive_network
                )
                st.session_state.negative_network_html = _build_interactive_network_html(
                    result.negative_network
                )
                st.session_state.layout_engine = result.layout_engine
                st.session_state.stats = result.stats

        except B3InputError as exc:
            _clear_report_state()
            st.error(str(exc))
        except Exception:
            _clear_report_state()
            st.error("Report se nepodařilo vygenerovat.")
            with st.expander("Technické detaily"):
                st.code(traceback.format_exc())


def _render_result() -> None:
    st.success("Report je připravený.")

    download_col, reset_col = st.columns([1, 1])
    with download_col:
        st.download_button(
            "Stáhnout report",
            data=st.session_state.report_bytes,
            file_name="B3_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    with reset_col:
        if st.button("Smazat report"):
            _reset_app()

    report_tab, network_tab = st.tabs(["Náhled reportu", "Interaktivní grafy"])

    with report_tab:
        preview_html = st.session_state.get("preview_html")
        if preview_html:
            components.html(_wrap_preview_html(preview_html), height=900, scrolling=True)
        else:
            st.info("Náhled není k dispozici, report lze stáhnout jako DOCX.")

    with network_tab:
        st.caption("Najetím na uzel nebo kliknutím zvýrazníte jeho příchozí a odchozí vazby.")
        positive_tab, negative_tab = st.tabs(["Kladné vztahy", "Záporné vztahy"])
        with positive_tab:
            components.html(
                st.session_state.positive_network_html,
                height=760,
                scrolling=False,
            )
        with negative_tab:
            components.html(
                st.session_state.negative_network_html,
                height=760,
                scrolling=False,
            )


def _clear_report_state() -> None:
    for key in REPORT_STATE_KEYS:
        st.session_state.pop(key, None)


def _reset_app() -> None:
    _clear_report_state()
    st.session_state.upload_key += 1
    st.rerun()


def _build_preview_html(result) -> str:
    stats = result.stats
    positive_chart = _image_data_uri(result.positive_chart_path)
    negative_chart = _image_data_uri(result.negative_chart_path)
    positive_table = result.positive_table.to_html(index=False, border=0)
    negative_table = result.negative_table.to_html(index=False, border=0)

    return f"""
      <h1>Grafický výstup z dotazníku B-3</h1>
      <p>Dva níže uvedené grafy zobrazují vztahy mezi členy třídního kolektivu
      na základě jejich pozitivních a negativních voleb týkajících se toho,
      koho ze svých spolužáků považují za své přátele.</p>

      <p><strong>Celkový počet žáků:</strong> {stats["total_students"]}<br>
      <strong>Počet chlapců a dívek:</strong> {stats["male_students"]} / {stats["female_students"]}<br>
      <strong>Počet nepřítomných žáků nebo bez souhlasu:</strong> {stats["not_present_or_without_consent"]}</p>

      <p><strong>Vysvětlivky ke grafům:</strong><br>
      &bull; Směr šipek reprezentuje, kdo koho označil ve své pozitivní nebo negativní volbě.<br>
      &bull; Barva uzlů odpovídá pohlaví žáků - žlutá barva reprezentuje dívky a fialová barva chlapce.<br>
      &bull; Uzly s červeným okrajem odpovídají žákům, kteří nebyli v době sběru dat přítomni nebo u nich není k dispozici souhlas se zpracováním jejich odpovědí.<br>
      &bull; Spodní číslo v uzlu odpovídá kódovému označení žáka.<br>
      &bull; Horní číslo v uzlu odpovídá součtu vážených odpovědí na otázky č. 1., resp. 2, a 6.<br>
      &bull; Vzdálenosti mezi uzly jsou zvoleny tak, aby co nejlépe odrážely vnitřní strukturu třídního kolektivu, tzn. tak, aby žáci s podobnými a vzájemnýmmi volbami byli v grafu blízko sebe.</p>

      <img src="{positive_chart}" alt="Kladné vztahy">
      <img src="{negative_chart}" alt="Záporné vztahy">

      <p><strong>Žáci sestupně seřazení podle počtu obdržených kladných bodů</strong></p>
      {positive_table}

      <p><strong>Žáci sestupně seřazení podle počtu obdržených záporných bodů</strong></p>
      {negative_table}
    """


def _image_data_uri(path: Path) -> str:
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def _build_interactive_network_html(network: dict) -> str:
    chart = _scale_network_for_svg(network)
    payload = json.dumps(chart, ensure_ascii=False)
    safe_title = html.escape(network["title"])

    return f"""
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          * {{
            box-sizing: border-box;
          }}
          body {{
            margin: 0;
            color: #151515;
            background: #ffffff;
            font-family: Calibri, Arial, sans-serif;
          }}
          .toolbar {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            min-height: 48px;
            padding: 0 4px 12px;
          }}
          .title {{
            margin: 0;
            font-size: 24px;
            font-weight: 700;
          }}
          .hint {{
            color: #555;
            font-size: 14px;
          }}
          button {{
            border: 1px solid #999;
            border-radius: 8px;
            padding: 7px 12px;
            color: #151515;
            background: #ffffff;
            cursor: pointer;
          }}
          .chart-wrap {{
            position: relative;
          }}
          svg {{
            display: block;
            width: 100%;
            height: 690px;
            border: 1px solid #d7d7d7;
            border-radius: 8px;
            background: #ffffff;
          }}
          .edge {{
            stroke: #7c7c7c;
            stroke-width: 1.7;
            opacity: 0.58;
            transition: opacity 120ms ease, stroke-width 120ms ease, stroke 120ms ease;
          }}
          .node circle {{
            stroke-width: 2.5;
            opacity: 0.72;
            cursor: pointer;
            transition: opacity 120ms ease, stroke-width 120ms ease, filter 120ms ease;
          }}
          .node text {{
            fill: #000000;
            font-size: 15px;
            text-anchor: middle;
            pointer-events: none;
            dominant-baseline: central;
          }}
          .dim {{
            opacity: 0.12;
          }}
          .edge.highlight-out {{
            stroke: #2d7dd2;
            stroke-width: 3.5;
            opacity: 1;
          }}
          .edge.highlight-in {{
            stroke: #d95f02;
            stroke-width: 3.5;
            opacity: 1;
          }}
          .node.highlight circle {{
            opacity: 1;
            stroke-width: 4;
            filter: drop-shadow(0 2px 5px rgba(0, 0, 0, 0.25));
          }}
          .node.selected circle {{
            stroke: #111111;
            stroke-width: 5;
            opacity: 1;
          }}
          .node.out-neighbor circle {{
            stroke: #2d7dd2;
            opacity: 1;
          }}
          .node.in-neighbor circle {{
            stroke: #d95f02;
            opacity: 1;
          }}
          .legend {{
            position: absolute;
            top: 12px;
            left: 12px;
            display: flex;
            flex-wrap: wrap;
            gap: 16px;
            padding: 8px 10px;
            color: #444;
            font-size: 14px;
            border: 1px solid #d7d7d7;
            border-radius: 8px;
            background: rgba(255, 255, 255, 0.92);
          }}
          .swatch {{
            display: inline-block;
            width: 20px;
            height: 3px;
            margin-right: 6px;
            vertical-align: middle;
          }}
          .out {{
            background: #2d7dd2;
          }}
          .in {{
            background: #d95f02;
          }}
        </style>
      </head>
      <body>
        <div class="toolbar">
          <h2 class="title">{safe_title}</h2>
          <span class="hint">Kliknutí uzel uzamkne, druhé kliknutí výběr zruší.</span>
          <button id="reset">Zrušit výběr</button>
        </div>
        <div class="chart-wrap">
          <svg id="network" viewBox="0 0 1000 650" aria-label="{safe_title}">
            <defs>
              <marker id="arrow" viewBox="0 0 10 10" refX="9" refY="5"
                markerWidth="7" markerHeight="7" orient="auto-start-reverse">
                <path d="M 0 0 L 10 5 L 0 10 z" fill="#7c7c7c"></path>
              </marker>
              <marker id="arrow-out" viewBox="0 0 10 10" refX="9" refY="5"
                markerWidth="8" markerHeight="8" orient="auto-start-reverse">
                <path d="M 0 0 L 10 5 L 0 10 z" fill="#2d7dd2"></path>
              </marker>
              <marker id="arrow-in" viewBox="0 0 10 10" refX="9" refY="5"
                markerWidth="8" markerHeight="8" orient="auto-start-reverse">
                <path d="M 0 0 L 10 5 L 0 10 z" fill="#d95f02"></path>
              </marker>
            </defs>
            <g id="edges"></g>
            <g id="nodes"></g>
          </svg>
          <div class="legend">
            <span><span class="swatch out"></span>Odchozí vazby</span>
            <span><span class="swatch in"></span>Příchozí vazby</span>
          </div>
        </div>
        <script>
          const data = {payload};
          const svg = document.getElementById("network");
          const edgeLayer = document.getElementById("edges");
          const nodeLayer = document.getElementById("nodes");
          const reset = document.getElementById("reset");
          const radius = 22;
          let pinnedNodeId = null;

          function edgeEndpoint(source, target) {{
            const dx = target.x - source.x;
            const dy = target.y - source.y;
            const length = Math.hypot(dx, dy) || 1;
            const ux = dx / length;
            const uy = dy / length;
            return {{
              x1: source.x + ux * radius,
              y1: source.y + uy * radius,
              x2: target.x - ux * (radius + 3),
              y2: target.y - uy * (radius + 3)
            }};
          }}

          function draw() {{
            const nodeById = new Map(data.nodes.map(node => [node.id, node]));

            data.edges.forEach((edge, index) => {{
              const source = nodeById.get(edge.source);
              const target = nodeById.get(edge.target);
              const line = edgeEndpoint(source, target);
              const element = document.createElementNS("http://www.w3.org/2000/svg", "line");
              element.setAttribute("x1", line.x1);
              element.setAttribute("y1", line.y1);
              element.setAttribute("x2", line.x2);
              element.setAttribute("y2", line.y2);
              element.setAttribute("marker-end", "url(#arrow)");
              element.classList.add("edge");
              element.dataset.source = edge.source;
              element.dataset.target = edge.target;
              element.dataset.index = String(index);
              edgeLayer.appendChild(element);
            }});

            data.nodes.forEach(node => {{
              const group = document.createElementNS("http://www.w3.org/2000/svg", "g");
              group.classList.add("node");
              group.dataset.id = node.id;
              group.setAttribute("transform", `translate(${{node.x}}, ${{node.y}})`);

              const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
              circle.setAttribute("r", radius);
              circle.setAttribute("fill", node.gender === "male" ? "#6e6bff" : "#fbff63");
              circle.setAttribute("stroke", node.presentWithConsent === "yes" ? "#a8a8a8" : "#ff6363");

              const score = document.createElementNS("http://www.w3.org/2000/svg", "text");
              score.setAttribute("y", "-8");
              score.textContent = String(node.pointsOverall);

              const label = document.createElementNS("http://www.w3.org/2000/svg", "text");
              label.setAttribute("y", "10");
              label.textContent = node.label;

              const title = document.createElementNS("http://www.w3.org/2000/svg", "title");
              title.textContent = `Žák ${{node.label}} | body: ${{node.pointsOverall}}`;

              group.appendChild(title);
              group.appendChild(circle);
              group.appendChild(score);
              group.appendChild(label);
              group.addEventListener("mouseenter", () => {{
                if (!pinnedNodeId) highlight(node.id);
              }});
              group.addEventListener("mouseleave", () => {{
                if (!pinnedNodeId) clearHighlight();
              }});
              group.addEventListener("click", () => {{
                pinnedNodeId = pinnedNodeId === node.id ? null : node.id;
                if (pinnedNodeId) highlight(pinnedNodeId);
                else clearHighlight();
              }});
              nodeLayer.appendChild(group);
            }});
          }}

          function highlight(nodeId) {{
            const incoming = new Set();
            const outgoing = new Set();
            const relevantEdges = new Set();

            document.querySelectorAll(".edge").forEach(edge => {{
              edge.classList.remove("highlight-in", "highlight-out", "dim");
              edge.setAttribute("marker-end", "url(#arrow)");
              if (edge.dataset.source === nodeId) {{
                outgoing.add(edge.dataset.target);
                relevantEdges.add(edge.dataset.index);
                edge.classList.add("highlight-out");
                edge.setAttribute("marker-end", "url(#arrow-out)");
              }} else if (edge.dataset.target === nodeId) {{
                incoming.add(edge.dataset.source);
                relevantEdges.add(edge.dataset.index);
                edge.classList.add("highlight-in");
                edge.setAttribute("marker-end", "url(#arrow-in)");
              }} else {{
                edge.classList.add("dim");
              }}
            }});

            document.querySelectorAll(".node").forEach(node => {{
              const id = node.dataset.id;
              node.classList.remove("highlight", "selected", "in-neighbor", "out-neighbor", "dim");
              if (id === nodeId) {{
                node.classList.add("highlight", "selected");
              }} else if (outgoing.has(id)) {{
                node.classList.add("highlight", "out-neighbor");
              }} else if (incoming.has(id)) {{
                node.classList.add("highlight", "in-neighbor");
              }} else {{
                node.classList.add("dim");
              }}
            }});
          }}

          function clearHighlight() {{
            document.querySelectorAll(".edge").forEach(edge => {{
              edge.classList.remove("highlight-in", "highlight-out", "dim");
              edge.setAttribute("marker-end", "url(#arrow)");
            }});
            document.querySelectorAll(".node").forEach(node => {{
              node.classList.remove("highlight", "selected", "in-neighbor", "out-neighbor", "dim");
            }});
          }}

          reset.addEventListener("click", () => {{
            pinnedNodeId = null;
            clearHighlight();
          }});

          svg.addEventListener("click", event => {{
            if (event.target === svg) {{
              pinnedNodeId = null;
              clearHighlight();
            }}
          }});

          draw();
        </script>
      </body>
    </html>
    """


def _scale_network_for_svg(network: dict) -> dict:
    width = 1000
    height = 650
    margin = 54
    nodes = network["nodes"]
    xs = [node["x"] for node in nodes]
    ys = [node["y"] for node in nodes]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    span_x = max(max_x - min_x, 1)
    span_y = max(max_y - min_y, 1)
    scale = min((width - 2 * margin) / span_x, (height - 2 * margin) / span_y)
    offset_x = (width - span_x * scale) / 2
    offset_y = (height - span_y * scale) / 2

    scaled_nodes = []
    for node in nodes:
        scaled = dict(node)
        scaled["x"] = offset_x + (node["x"] - min_x) * scale
        scaled["y"] = height - (offset_y + (node["y"] - min_y) * scale)
        scaled_nodes.append(scaled)

    return {
        "title": network["title"],
        "nodes": scaled_nodes,
        "edges": network["edges"],
    }


def _wrap_preview_html(body: str) -> str:
    return f"""
    <!doctype html>
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          body {{
            box-sizing: border-box;
            max-width: 900px;
            margin: 0 auto;
            padding: 32px;
            color: #151515;
            font-family: Calibri, Arial, sans-serif;
            font-size: 16px;
            line-height: 1.45;
            background: #ffffff;
          }}
          img {{
            display: block;
            max-width: 100%;
            height: auto;
            margin: 22px 0;
          }}
          table {{
            border-collapse: collapse;
            width: 100%;
            margin: 14px 0 28px;
          }}
          th, td {{
            border: 1px solid #777;
            padding: 6px 8px;
            text-align: left;
          }}
        </style>
      </head>
      <body>{body}</body>
    </html>
    """


if __name__ == "__main__":
    main()
