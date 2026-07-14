#!/usr/bin/env python3
"""Render architecture and workflow diagrams from product-agnostic JSON.

The script intentionally accepts explicit nodes and edges instead of inferring
software behavior.  This keeps generated copyright figures tied to evidence
assembled elsewhere in the skill workflow.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw, ImageFont


DEFAULT_PALETTE = [
    ("#EAF3FF", "#2F79DA"),
    ("#ECF8EE", "#46A15B"),
    ("#FFF4E8", "#E48A2F"),
    ("#F4EFFF", "#8A66C7"),
    ("#FFF9E6", "#D3A428"),
    ("#EAF8F8", "#2A9292"),
]
FONT_CANDIDATES = [
    Path("C:/Windows/Fonts/msyh.ttc"),
    Path("C:/Windows/Fonts/simhei.ttf"),
    Path("C:/Windows/Fonts/simsun.ttc"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"),
    Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"),
    Path("/System/Library/Fonts/PingFang.ttc"),
    Path("/System/Library/Fonts/STHeiti Medium.ttc"),
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
]
BOLD_FONT_CANDIDATES = [
    Path("C:/Windows/Fonts/msyhbd.ttc"),
    Path("C:/Windows/Fonts/simhei.ttf"),
    Path("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
    Path("/usr/share/fonts/truetype/noto/NotoSansCJK-Bold.ttc"),
    Path("/System/Library/Fonts/PingFang.ttc"),
    Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
]


class DiagramConfigError(ValueError):
    """Raised when a diagram specification is invalid."""


@dataclass(frozen=True)
class Box:
    left: float
    top: float
    right: float
    bottom: float

    @property
    def center(self) -> tuple[float, float]:
        return ((self.left + self.right) / 2, (self.top + self.bottom) / 2)

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.bottom - self.top

    def as_tuple(self) -> tuple[int, int, int, int]:
        return tuple(round(value) for value in (self.left, self.top, self.right, self.bottom))  # type: ignore[return-value]


def _mapping(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise DiagramConfigError(f"{label} must be a JSON object")
    return value


def _list(value: Any, label: str) -> list[Any]:
    if value is None:
        return []
    if not isinstance(value, list):
        raise DiagramConfigError(f"{label} must be a JSON array")
    return value


def _text(value: Any) -> str:
    return "" if value is None else str(value)


def _int(value: Any, default: int, label: str) -> int:
    if value is None:
        return default
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise DiagramConfigError(f"{label} must be an integer") from exc


def _float(value: Any, default: float, label: str) -> float:
    if value is None:
        return default
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise DiagramConfigError(f"{label} must be a number") from exc


def color(value: Any, default: str) -> str:
    raw = _text(value).strip() or default
    if not re.fullmatch(r"#[0-9A-Fa-f]{6}", raw):
        raise DiagramConfigError(f"invalid color: {raw!r}; use #RRGGBB")
    return raw.upper()


def find_font(explicit: Path | None, candidates: list[Path], label: str) -> Path:
    if explicit is not None:
        path = explicit.expanduser().resolve()
        if not path.exists():
            raise FileNotFoundError(path)
        return path
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise DiagramConfigError(
        f"no usable {label} font found; pass --font/--bold-font with a Unicode TrueType or OpenType font"
    )


class FontBook:
    def __init__(self, regular: Path, bold: Path) -> None:
        self.regular_path = regular
        self.bold_path = bold
        self._cache: dict[tuple[int, bool], ImageFont.FreeTypeFont] = {}

    def get(self, size: int, *, bold: bool = False) -> ImageFont.FreeTypeFont:
        size = max(8, int(size))
        key = (size, bold)
        if key not in self._cache:
            self._cache[key] = ImageFont.truetype(str(self.bold_path if bold else self.regular_path), size=size)
        return self._cache[key]


def make_background(width: int, height: int, top: str, bottom: str) -> Image.Image:
    image = Image.new("RGB", (width, height), top)
    draw = ImageDraw.Draw(image)
    top_rgb = tuple(int(top[index : index + 2], 16) for index in (1, 3, 5))
    bottom_rgb = tuple(int(bottom[index : index + 2], 16) for index in (1, 3, 5))
    for y in range(height):
        ratio = y / max(1, height - 1)
        rgb = tuple(round(top_rgb[i] * (1 - ratio) + bottom_rgb[i] * ratio) for i in range(3))
        draw.line((0, y, width, y), fill=rgb)
    return image


def text_size(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont) -> tuple[int, int]:
    if not text:
        bbox = draw.textbbox((0, 0), " ", font=font)
    else:
        bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    if max_width < 10:
        return [text]
    output: list[str] = []
    for raw_line in text.splitlines() or [""]:
        if not raw_line:
            output.append("")
            continue
        current = ""
        for char in raw_line:
            trial = current + char
            if not current or text_size(draw, trial, font)[0] <= max_width:
                current = trial
            else:
                output.append(current)
                current = char
        if current:
            output.append(current)
    return output or [""]


def draw_centered_text(
    draw: ImageDraw.ImageDraw,
    box: Box,
    text: str,
    font: ImageFont.FreeTypeFont,
    *,
    fill: str = "#17324D",
    padding: int = 20,
    line_gap: int = 7,
) -> None:
    lines = wrap_text(draw, text, font, max(10, round(box.width) - padding * 2))
    dimensions = [text_size(draw, line, font) for line in lines]
    total_height = sum(height for _, height in dimensions) + max(0, len(lines) - 1) * line_gap
    y = box.top + (box.height - total_height) / 2
    for line, (width, height) in zip(lines, dimensions):
        draw.text((box.left + (box.width - width) / 2, y), line, font=font, fill=fill)
        y += height + line_gap


def rounded_box(
    draw: ImageDraw.ImageDraw,
    box: Box,
    *,
    fill: str,
    outline: str,
    radius: int = 18,
    width: int = 2,
) -> None:
    draw.rounded_rectangle(box.as_tuple(), radius=radius, fill=fill, outline=outline, width=width)


def draw_dashed_segment(
    draw: ImageDraw.ImageDraw,
    start: tuple[float, float],
    end: tuple[float, float],
    *,
    fill: str,
    width: int,
    dash: float = 12,
    gap: float = 8,
) -> None:
    dx, dy = end[0] - start[0], end[1] - start[1]
    length = math.hypot(dx, dy)
    if length < 1:
        return
    ux, uy = dx / length, dy / length
    position = 0.0
    while position < length:
        segment_end = min(length, position + dash)
        draw.line(
            (
                start[0] + ux * position,
                start[1] + uy * position,
                start[0] + ux * segment_end,
                start[1] + uy * segment_end,
            ),
            fill=fill,
            width=width,
        )
        position += dash + gap


def arrow_head(
    draw: ImageDraw.ImageDraw,
    point: tuple[float, float],
    from_point: tuple[float, float],
    *,
    fill: str,
    size: float = 15,
) -> None:
    dx, dy = point[0] - from_point[0], point[1] - from_point[1]
    length = max(1.0, math.hypot(dx, dy))
    ux, uy = dx / length, dy / length
    px, py = -uy, ux
    back_x, back_y = point[0] - ux * size, point[1] - uy * size
    draw.polygon(
        [
            point,
            (back_x + px * size * 0.55, back_y + py * size * 0.55),
            (back_x - px * size * 0.55, back_y - py * size * 0.55),
        ],
        fill=fill,
    )


def polyline_midpoint(points: list[tuple[float, float]]) -> tuple[float, float]:
    lengths = [math.dist(a, b) for a, b in zip(points, points[1:])]
    total = sum(lengths)
    if total <= 0:
        return points[0]
    target = total / 2
    traversed = 0.0
    for (a, b), length in zip(zip(points, points[1:]), lengths):
        if traversed + length >= target:
            ratio = (target - traversed) / max(length, 1)
            return (a[0] + (b[0] - a[0]) * ratio, a[1] + (b[1] - a[1]) * ratio)
        traversed += length
    return points[-1]


def draw_arrow(
    draw: ImageDraw.ImageDraw,
    points: list[tuple[float, float]],
    *,
    fill: str,
    width: int = 4,
    dashed: bool = False,
    bidirectional: bool = False,
    label: str = "",
    label_font: ImageFont.FreeTypeFont | None = None,
    label_fill: str = "#40576E",
) -> None:
    if len(points) < 2:
        raise DiagramConfigError("an edge needs at least two points")
    for start, end in zip(points, points[1:]):
        if dashed:
            draw_dashed_segment(draw, start, end, fill=fill, width=width)
        else:
            draw.line((start[0], start[1], end[0], end[1]), fill=fill, width=width)
    arrow_head(draw, points[-1], points[-2], fill=fill)
    if bidirectional:
        arrow_head(draw, points[0], points[1], fill=fill)
    if label and label_font is not None:
        x, y = polyline_midpoint(points)
        tw, th = text_size(draw, label, label_font)
        label_box = (round(x - tw / 2 - 8), round(y - th / 2 - 5), round(x + tw / 2 + 8), round(y + th / 2 + 5))
        draw.rounded_rectangle(label_box, radius=6, fill="#FFFFFF", outline="#D5DEE7", width=1)
        draw.text((x - tw / 2, y - th / 2), label, font=label_font, fill=label_fill)


def node_id(raw: Any, fallback: str) -> str:
    if isinstance(raw, str):
        return fallback
    if not isinstance(raw, dict):
        raise DiagramConfigError("node must be a string or object")
    value = _text(raw.get("id")).strip()
    return value or fallback


def node_label(raw: Any) -> str:
    if isinstance(raw, str):
        return raw
    value = _mapping(raw, "node")
    label = _text(value.get("label") or value.get("title")).strip()
    subtitle = _text(value.get("subtitle")).strip()
    return f"{label}\n{subtitle}" if label and subtitle else label or subtitle


def connection_points(source: Box, target: Box, *, orthogonal: bool = True) -> list[tuple[float, float]]:
    sx, sy = source.center
    tx, ty = target.center
    dx, dy = tx - sx, ty - sy
    if abs(dx) >= abs(dy):
        if dx >= 0:
            start = (source.right, sy)
            end = (target.left, ty)
        else:
            start = (source.left, sy)
            end = (target.right, ty)
        if orthogonal and abs(sy - ty) > 2:
            mid_x = (start[0] + end[0]) / 2
            return [start, (mid_x, start[1]), (mid_x, end[1]), end]
        return [start, end]
    if dy >= 0:
        start = (sx, source.bottom)
        end = (tx, target.top)
    else:
        start = (sx, source.top)
        end = (tx, target.bottom)
    if orthogonal and abs(sx - tx) > 2:
        mid_y = (start[1] + end[1]) / 2
        return [start, (start[0], mid_y), (end[0], mid_y), end]
    return [start, end]


def parse_edge(raw: Any, boxes: dict[str, Box], label: str) -> tuple[dict[str, Any], list[tuple[float, float]]]:
    edge = _mapping(raw, label)
    source_id = _text(edge.get("from")).strip()
    target_id = _text(edge.get("to")).strip()
    if source_id not in boxes:
        raise DiagramConfigError(f"edge references unknown source node: {source_id!r}")
    if target_id not in boxes:
        raise DiagramConfigError(f"edge references unknown target node: {target_id!r}")
    return edge, connection_points(boxes[source_id], boxes[target_id], orthogonal=bool(edge.get("orthogonal", True)))


def architecture_canvas_size(spec: dict[str, Any]) -> tuple[int, int]:
    layers = _list(spec.get("layers"), "architecture.layers")
    width = _int(spec.get("width"), 1800, "architecture.width")
    height = _int(spec.get("height"), 1050, "architecture.height")
    max_nodes = 1
    for raw_layer in layers:
        layer = _mapping(raw_layer, "architecture layer")
        max_nodes = max(max_nodes, len(_list(layer.get("nodes"), "architecture layer.nodes")))
    if bool(spec.get("auto_expand", True)):
        width = max(width, 280 + max_nodes * 220)
        height = max(height, 120 + len(layers) * 175)
    if width < 800 or height < 500:
        raise DiagramConfigError("architecture canvas must be at least 800x500")
    return width, height


def render_architecture(spec: dict[str, Any], fonts: FontBook) -> Image.Image:
    layers_raw = _list(spec.get("layers"), "architecture.layers")
    if not layers_raw:
        raise DiagramConfigError("architecture.layers cannot be empty")
    width, height = architecture_canvas_size(spec)
    image = make_background(
        width,
        height,
        color(spec.get("background_top"), "#FFFFFF"),
        color(spec.get("background_bottom"), "#F4F9FF"),
    )
    draw = ImageDraw.Draw(image)
    margin = _int(spec.get("margin"), 24, "architecture.margin")
    title = _text(spec.get("title")).strip()
    title_height = _int(spec.get("title_height"), 78 if title else 0, "architecture.title_height")
    if title:
        title_font = fonts.get(_int(spec.get("title_font_size"), 34, "architecture.title_font_size"), bold=True)
        tw, th = text_size(draw, title, title_font)
        draw.text(((width - tw) / 2, max(12, (title_height - th) / 2)), title, font=title_font, fill=color(spec.get("title_color"), "#17324D"))

    label_width = _int(spec.get("layer_label_width"), 150, "architecture.layer_label_width")
    layer_gap = _int(spec.get("layer_gap"), 16, "architecture.layer_gap")
    usable_height = height - title_height - margin * 2 - layer_gap * (len(layers_raw) - 1)
    layer_height = usable_height / len(layers_raw)
    if layer_height < 115:
        raise DiagramConfigError("architecture layers are too short; increase height or enable auto_expand")

    node_boxes: dict[str, Box] = {}
    layer_bounds: list[Box] = []
    layer_specs: list[dict[str, Any]] = []
    palette_raw = _list(spec.get("palette"), "architecture.palette")
    for layer_index, raw_layer in enumerate(layers_raw):
        layer = _mapping(raw_layer, f"architecture.layers[{layer_index}]")
        nodes = _list(layer.get("nodes"), f"architecture.layers[{layer_index}].nodes")
        if not nodes:
            raise DiagramConfigError(f"architecture layer {layer_index + 1} has no nodes")
        if palette_raw:
            pair_raw = palette_raw[layer_index % len(palette_raw)]
            if not isinstance(pair_raw, list) or len(pair_raw) != 2:
                raise DiagramConfigError("architecture.palette entries must be [fill, outline]")
            default_fill, default_outline = color(pair_raw[0], "#EAF3FF"), color(pair_raw[1], "#2F79DA")
        else:
            default_fill, default_outline = DEFAULT_PALETTE[layer_index % len(DEFAULT_PALETTE)]
        fill = color(layer.get("fill"), default_fill)
        outline = color(layer.get("outline"), default_outline)
        y1 = title_height + margin + layer_index * (layer_height + layer_gap)
        y2 = y1 + layer_height
        label_box = Box(margin, y1, margin + label_width, y2)
        content_box = Box(margin + label_width + 18, y1, width - margin, y2)
        layer_bounds.append(content_box)
        layer_specs.append(layer)
        rounded_box(draw, label_box, fill=fill, outline=outline, radius=18, width=2)
        draw_centered_text(
            draw,
            label_box,
            _text(layer.get("label") or layer.get("name")).strip() or f"Layer {layer_index + 1}",
            fonts.get(_int(layer.get("label_font_size"), 27, "architecture layer.label_font_size"), bold=True),
            fill=outline,
            padding=15,
        )
        rounded_box(draw, content_box, fill=fill, outline=outline, radius=20, width=2)

        layer_title = _text(layer.get("title")).strip()
        title_band = 48 if layer_title else 8
        if layer_title:
            layer_title_font = fonts.get(_int(layer.get("title_font_size"), 24, "architecture layer.title_font_size"), bold=True)
            tw, _ = text_size(draw, layer_title, layer_title_font)
            draw.text(((content_box.left + content_box.right - tw) / 2, content_box.top + 9), layer_title, font=layer_title_font, fill="#1C2F42")

        inner_left = content_box.left + 24
        inner_right = content_box.right - 24
        node_gap = _int(layer.get("node_gap"), 14, "architecture layer.node_gap")
        node_width = (inner_right - inner_left - node_gap * (len(nodes) - 1)) / len(nodes)
        node_top = content_box.top + title_band + 8
        node_bottom = content_box.bottom - 18
        if node_width < 75 or node_bottom - node_top < 55:
            raise DiagramConfigError(f"architecture layer {layer_index + 1} nodes do not fit")
        for node_index, raw_node in enumerate(nodes):
            identifier = node_id(raw_node, f"layer-{layer_index + 1}-node-{node_index + 1}")
            if identifier in node_boxes:
                raise DiagramConfigError(f"duplicate node id: {identifier}")
            left = inner_left + node_index * (node_width + node_gap)
            box = Box(left, node_top, left + node_width, node_bottom)
            node_boxes[identifier] = box
            node_cfg = raw_node if isinstance(raw_node, dict) else {}
            node_fill = color(node_cfg.get("fill"), "#FFFFFF")
            node_outline = color(node_cfg.get("outline"), outline)
            rounded_box(draw, box, fill=node_fill, outline=node_outline, radius=14, width=2)
            label = node_label(raw_node)
            if not label:
                raise DiagramConfigError(f"architecture node {identifier!r} has no label")
            draw_centered_text(
                draw,
                box,
                label,
                fonts.get(_int(node_cfg.get("font_size"), 20 if len(nodes) >= 6 else 22, "architecture node.font_size"), bold=bool(node_cfg.get("bold", False))),
                fill=color(node_cfg.get("text_color"), "#17324D"),
                padding=12,
                line_gap=5,
            )

    edges_raw = _list(spec.get("edges"), "architecture.edges")
    edge_font = fonts.get(_int(spec.get("edge_label_font_size"), 17, "architecture.edge_label_font_size"))
    if edges_raw:
        for edge_index, raw_edge in enumerate(edges_raw):
            edge, points = parse_edge(raw_edge, node_boxes, f"architecture.edges[{edge_index}]")
            draw_arrow(
                draw,
                points,
                fill=color(edge.get("color"), "#637D96"),
                width=_int(edge.get("width"), 3, "architecture edge.width"),
                dashed=bool(edge.get("dashed", False)),
                bidirectional=bool(edge.get("bidirectional", False)),
                label=_text(edge.get("label")).strip(),
                label_font=edge_font,
            )
    elif bool(spec.get("connect_adjacent_layers", True)) and len(layer_bounds) > 1:
        connector_count = max(1, _int(spec.get("layer_connector_count"), 3, "architecture.layer_connector_count"))
        for upper, lower in zip(layer_bounds, layer_bounds[1:]):
            for index in range(connector_count):
                ratio = (index + 1) / (connector_count + 1)
                x = upper.left + upper.width * ratio
                draw_arrow(
                    draw,
                    [(x, upper.bottom + 2), (x, lower.top - 2)],
                    fill=color(spec.get("connector_color"), "#73889B"),
                    width=3,
                    bidirectional=bool(spec.get("layer_connectors_bidirectional", True)),
                )
    return image


def flatten_flow_nodes(spec: dict[str, Any], lanes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    nodes: list[dict[str, Any]] = []
    for lane_index, lane in enumerate(lanes):
        lane_id = _text(lane.get("id")).strip() or f"lane-{lane_index + 1}"
        for node_index, raw_node in enumerate(_list(lane.get("nodes"), f"flow lane {lane_id}.nodes")):
            if isinstance(raw_node, str):
                node = {"label": raw_node}
            else:
                node = dict(_mapping(raw_node, f"flow lane {lane_id} node"))
            node.setdefault("lane", lane_id)
            node.setdefault("column", node_index)
            node.setdefault("id", f"{lane_id}-node-{node_index + 1}")
            nodes.append(node)
    for node_index, raw_node in enumerate(_list(spec.get("nodes"), "flow.nodes")):
        node = dict(_mapping(raw_node, f"flow.nodes[{node_index}]"))
        node.setdefault("id", f"node-{node_index + 1}")
        node.setdefault("column", node_index)
        nodes.append(node)
    return nodes


def flow_canvas_size(spec: dict[str, Any], lane_count: int, column_count: int) -> tuple[int, int]:
    width = _int(spec.get("width"), 1800, "flow.width")
    height = _int(spec.get("height"), 980, "flow.height")
    if bool(spec.get("auto_expand", True)):
        width = max(width, 250 + column_count * 210)
        height = max(height, 120 + lane_count * 215)
    if width < 800 or height < 500:
        raise DiagramConfigError("flow canvas must be at least 800x500")
    return width, height


def draw_flow_node(
    draw: ImageDraw.ImageDraw,
    box: Box,
    node: dict[str, Any],
    fonts: FontBook,
    default_outline: str,
) -> None:
    shape = _text(node.get("shape")).strip().lower() or "action"
    fill = color(node.get("fill"), "#FFFFFF")
    outline = color(node.get("outline"), default_outline)
    if shape in {"terminal", "start", "end"}:
        draw.rounded_rectangle(box.as_tuple(), radius=round(box.height / 2), fill=fill, outline=outline, width=3)
    elif shape == "decision":
        cx, cy = box.center
        points = [(cx, box.top), (box.right, cy), (cx, box.bottom), (box.left, cy)]
        draw.polygon(points, fill=fill)
        draw.line(points + [points[0]], fill=outline, width=3)
    elif shape == "data":
        skew = min(24, box.width * 0.12)
        points = [(box.left + skew, box.top), (box.right, box.top), (box.right - skew, box.bottom), (box.left, box.bottom)]
        draw.polygon(points, fill=fill)
        draw.line(points + [points[0]], fill=outline, width=3)
    elif shape == "action":
        rounded_box(draw, box, fill=fill, outline=outline, radius=16, width=2)
    else:
        raise DiagramConfigError(f"unsupported flow node shape: {shape}")
    label = _text(node.get("label") or node.get("title")).strip()
    if not label:
        raise DiagramConfigError(f"flow node {_text(node.get('id'))!r} has no label")
    draw_centered_text(
        draw,
        box,
        label,
        fonts.get(_int(node.get("font_size"), 21, "flow node.font_size"), bold=bool(node.get("bold", False))),
        fill=color(node.get("text_color"), "#18334E"),
        padding=18 if shape != "decision" else 34,
        line_gap=5,
    )


def render_flow(spec: dict[str, Any], fonts: FontBook) -> Image.Image:
    lanes_raw = _list(spec.get("lanes"), "flow.lanes")
    if not lanes_raw:
        raise DiagramConfigError("flow.lanes cannot be empty")
    lanes: list[dict[str, Any]] = []
    lane_ids: set[str] = set()
    for index, raw_lane in enumerate(lanes_raw):
        lane = dict(_mapping(raw_lane, f"flow.lanes[{index}]"))
        identifier = _text(lane.get("id")).strip() or f"lane-{index + 1}"
        if identifier in lane_ids:
            raise DiagramConfigError(f"duplicate lane id: {identifier}")
        lane["id"] = identifier
        lane_ids.add(identifier)
        lanes.append(lane)

    nodes = flatten_flow_nodes(spec, lanes)
    if not nodes:
        raise DiagramConfigError("flow must contain nodes")
    seen_nodes: set[str] = set()
    max_column = 0
    occupied: set[tuple[str, int]] = set()
    for index, node in enumerate(nodes):
        identifier = _text(node.get("id")).strip() or f"node-{index + 1}"
        if identifier in seen_nodes:
            raise DiagramConfigError(f"duplicate node id: {identifier}")
        node["id"] = identifier
        seen_nodes.add(identifier)
        lane_id = _text(node.get("lane")).strip()
        if lane_id not in lane_ids:
            raise DiagramConfigError(f"node {identifier!r} references unknown lane {lane_id!r}")
        column = _int(node.get("column"), index, "flow node.column")
        if column < 0:
            raise DiagramConfigError("flow node.column cannot be negative")
        node["column"] = column
        slot = (lane_id, column)
        if slot in occupied:
            raise DiagramConfigError(f"multiple nodes occupy lane/column slot {slot}")
        occupied.add(slot)
        max_column = max(max_column, column)

    column_count = max_column + 1
    width, height = flow_canvas_size(spec, len(lanes), column_count)
    image = make_background(
        width,
        height,
        color(spec.get("background_top"), "#FFFFFF"),
        color(spec.get("background_bottom"), "#F7FBFF"),
    )
    draw = ImageDraw.Draw(image)
    margin = _int(spec.get("margin"), 28, "flow.margin")
    title = _text(spec.get("title")).strip()
    title_height = _int(spec.get("title_height"), 78 if title else 0, "flow.title_height")
    if title:
        title_font = fonts.get(_int(spec.get("title_font_size"), 34, "flow.title_font_size"), bold=True)
        tw, th = text_size(draw, title, title_font)
        draw.text(((width - tw) / 2, max(12, (title_height - th) / 2)), title, font=title_font, fill=color(spec.get("title_color"), "#17324D"))

    lane_label_width = _int(spec.get("lane_label_width"), 180, "flow.lane_label_width")
    lane_gap = _int(spec.get("lane_gap"), 20, "flow.lane_gap")
    usable_height = height - title_height - margin * 2 - lane_gap * (len(lanes) - 1)
    lane_height = usable_height / len(lanes)
    if lane_height < 130:
        raise DiagramConfigError("flow lanes are too short; increase height or enable auto_expand")
    content_left = margin + lane_label_width
    content_right = width - margin
    column_width = (content_right - content_left) / column_count
    if column_width < 100:
        raise DiagramConfigError("flow columns are too narrow; increase width or enable auto_expand")

    lane_boxes: dict[str, Box] = {}
    lane_outlines: dict[str, str] = {}
    for lane_index, lane in enumerate(lanes):
        default_fill, default_outline = DEFAULT_PALETTE[lane_index % len(DEFAULT_PALETTE)]
        fill = color(lane.get("fill"), default_fill)
        outline = color(lane.get("outline"), default_outline)
        y1 = title_height + margin + lane_index * (lane_height + lane_gap)
        y2 = y1 + lane_height
        lane_box = Box(margin, y1, width - margin, y2)
        lane_boxes[lane["id"]] = lane_box
        lane_outlines[lane["id"]] = outline
        rounded_box(draw, lane_box, fill=fill, outline=outline, radius=22, width=2)
        label_box = Box(margin + 10, y1 + 8, margin + lane_label_width - 10, y2 - 8)
        draw_centered_text(
            draw,
            label_box,
            _text(lane.get("label") or lane.get("name")).strip() or lane["id"],
            fonts.get(_int(lane.get("font_size"), 25, "flow lane.font_size"), bold=True),
            fill=outline,
            padding=12,
        )
        draw.line((content_left - 12, y1 + 14, content_left - 12, y2 - 14), fill=outline, width=2)

    boxes: dict[str, Box] = {}
    node_height_default = _int(spec.get("node_height"), 104, "flow.node_height")
    for node in nodes:
        lane_box = lane_boxes[_text(node.get("lane"))]
        column = _int(node.get("column"), 0, "flow node.column")
        node_width = min(
            _float(node.get("width"), 250.0, "flow node.width"),
            max(78.0, column_width * _float(node.get("width_ratio"), 0.76, "flow node.width_ratio")),
        )
        node_height = _float(node.get("height"), node_height_default, "flow node.height")
        if _text(node.get("shape")).strip().lower() == "decision":
            node_height = max(node_height, 120)
        cx = content_left + (column + 0.5) * column_width
        cy = (lane_box.top + lane_box.bottom) / 2 + _float(node.get("offset_y"), 0.0, "flow node.offset_y")
        box = Box(cx - node_width / 2, cy - node_height / 2, cx + node_width / 2, cy + node_height / 2)
        if box.top < lane_box.top + 12 or box.bottom > lane_box.bottom - 12:
            raise DiagramConfigError(f"flow node {node['id']!r} does not fit its lane")
        boxes[node["id"]] = box

    edges_raw = _list(spec.get("edges"), "flow.edges")
    if not edges_raw and bool(spec.get("auto_connect", True)):
        ordered = sorted(nodes, key=lambda item: (_int(item.get("column"), 0, "flow node.column"), lanes.index(next(lane for lane in lanes if lane["id"] == item["lane"]))))
        edges_raw = [{"from": left["id"], "to": right["id"]} for left, right in zip(ordered, ordered[1:])]
    edge_font = fonts.get(_int(spec.get("edge_label_font_size"), 17, "flow.edge_label_font_size"))
    for edge_index, raw_edge in enumerate(edges_raw):
        edge, points = parse_edge(raw_edge, boxes, f"flow.edges[{edge_index}]")
        draw_arrow(
            draw,
            points,
            fill=color(edge.get("color"), "#637D96"),
            width=_int(edge.get("width"), 4, "flow edge.width"),
            dashed=bool(edge.get("dashed", False)),
            bidirectional=bool(edge.get("bidirectional", False)),
            label=_text(edge.get("label")).strip(),
            label_font=edge_font,
            label_fill=color(edge.get("label_color"), "#40576E"),
        )

    for node in nodes:
        lane_id = _text(node.get("lane"))
        draw_flow_node(draw, boxes[node["id"]], node, fonts, lane_outlines[lane_id])
    return image


def diagram_specs(payload: dict[str, Any]) -> list[dict[str, Any]]:
    global_style = _mapping(payload.get("style", {}), "style")
    raw_diagrams = _list(payload.get("diagrams"), "diagrams")
    collected: list[dict[str, Any]] = []
    if raw_diagrams:
        for index, raw in enumerate(raw_diagrams):
            spec = dict(global_style)
            spec.update(_mapping(raw, f"diagrams[{index}]"))
            collected.append(spec)
    else:
        for kind in ("architecture", "flow"):
            if kind in payload:
                spec = dict(global_style)
                spec.update(_mapping(payload[kind], kind))
                spec.setdefault("type", kind)
                collected.append(spec)
        if not collected and "type" in payload:
            spec = dict(global_style)
            spec.update(payload)
            collected.append(spec)
    if not collected:
        raise DiagramConfigError("JSON must contain diagrams[], architecture, flow, or a top-level type")
    return collected


def output_path(output_dir: Path, spec: dict[str, Any], kind: str, index: int) -> Path:
    default_name = f"{kind}.png" if index == 0 else f"{kind}-{index + 1}.png"
    raw = _text(spec.get("output")).strip() or default_name
    name = Path(raw).name
    if Path(name).suffix.lower() != ".png":
        name = f"{Path(name).stem}.png"
    return output_dir / name


def load_payload(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise DiagramConfigError(f"invalid JSON in {path}: {exc}") from exc
    return _mapping(value, "root")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    example = r'''
JSON example:
  {
    "architecture": {
      "title": "System Architecture", "output": "architecture.png",
      "layers": [
        {"id": "ui", "label": "Presentation", "title": "Client",
         "nodes": [{"id": "portal", "label": "User portal"}]},
        {"id": "service", "label": "Service", "title": "Business services",
         "nodes": [{"id": "core", "label": "Core processing"}]}
      ],
      "edges": [{"from": "portal", "to": "core", "bidirectional": true}]
    },
    "flow": {
      "title": "Core Workflow", "output": "workflow.png",
      "lanes": [{"id": "user", "label": "User"},
                {"id": "system", "label": "System"}],
      "nodes": [
        {"id": "start", "lane": "user", "column": 0,
         "label": "Start", "shape": "terminal"},
        {"id": "submit", "lane": "user", "column": 1,
         "label": "Submit data"},
        {"id": "process", "lane": "system", "column": 2,
         "label": "Process"},
        {"id": "end", "lane": "user", "column": 3,
         "label": "View result", "shape": "terminal"}
      ],
      "edges": [{"from": "start", "to": "submit"},
                {"from": "submit", "to": "process"},
                {"from": "process", "to": "end"}]
    }
  }

Node and lane labels may contain Chinese text. Use --font when the host does not
provide a CJK-capable system font. Output is always embedded RGB PNG at 150 dpi.
'''
    parser = argparse.ArgumentParser(
        description="Render architecture and workflow PNG diagrams from JSON nodes and edges.",
        epilog=example,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--input", required=True, type=Path, help="UTF-8 JSON diagram specification")
    parser.add_argument("--output-dir", required=True, type=Path, help="directory for generated PNG files")
    parser.add_argument("--only", choices=("all", "architecture", "flow"), default="all", help="render only one diagram type")
    parser.add_argument("--font", type=Path, help="regular Unicode TrueType/OpenType font")
    parser.add_argument("--bold-font", type=Path, help="bold Unicode TrueType/OpenType font (defaults to --font when supplied)")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    input_path = args.input.expanduser().resolve()
    output_dir = args.output_dir.expanduser().resolve()
    payload = load_payload(input_path)
    regular = find_font(args.font, FONT_CANDIDATES, "regular")
    explicit_bold = args.bold_font or args.font
    bold = find_font(explicit_bold, BOLD_FONT_CANDIDATES, "bold")
    fonts = FontBook(regular, bold)
    output_dir.mkdir(parents=True, exist_ok=True)

    outputs: list[dict[str, Any]] = []
    kind_counts = {"architecture": 0, "flow": 0}
    for spec in diagram_specs(payload):
        kind = _text(spec.get("type")).strip().lower()
        if kind not in kind_counts:
            raise DiagramConfigError(f"unsupported diagram type: {kind!r}")
        if args.only != "all" and args.only != kind:
            continue
        index = kind_counts[kind]
        kind_counts[kind] += 1
        image = render_architecture(spec, fonts) if kind == "architecture" else render_flow(spec, fonts)
        path = output_path(output_dir, spec, kind, index)
        image.save(path, format="PNG", optimize=True, dpi=(150, 150))
        outputs.append({"type": kind, "path": str(path), "width": image.width, "height": image.height})
    if not outputs:
        raise DiagramConfigError(f"no diagrams matched --only {args.only!r}")
    print(json.dumps({"outputs": outputs}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (DiagramConfigError, FileNotFoundError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        raise SystemExit(2)
