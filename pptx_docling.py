import json
import os
import sys
import glob as globmod

from docling.document_converter import DocumentConverter


def print_keys_hierarchical(data, indent=0):
    """
    מדפיס רק את המפתחות של ה-JSON בצורה היררכית
    """
    if isinstance(data, dict):
        for key, value in data.items():
            print("  " * indent + str(key))
            if isinstance(value, (dict, list)):
                print_keys_hierarchical(value, indent + 1)
    elif isinstance(data, list):
        if len(data) > 0:
            print("  " * indent + "[list item structure]:")
            print_keys_hierarchical(data[0], indent + 1)


def extract_reading_order(pptx_path, output_dir="reading_order_outputs"):
    """Extract and save the reading order of elements using Docling.

    Reading order for Docling: Docling converts documents through a pipeline that
    includes layout analysis and a reading-order model. The order of items in the
    exported document reflects Docling's inferred reading order, which may differ
    from the raw XML shape order. Docling groups content into semantic blocks
    (titles, text, tables, pictures, lists) and orders them based on its layout
    analysis. This is a DERIVED/INFERRED reading order, not raw XML order.
    """
    os.makedirs(output_dir, exist_ok=True)

    converter = DocumentConverter()
    doc = converter.convert(pptx_path)
    data = doc.document.export_to_dict()

    pptx_name = os.path.splitext(os.path.basename(pptx_path))[0]
    output_filename = f"docling__{pptx_name}.txt"
    output_path = os.path.join(output_dir, output_filename)

    lines = [
        f"Reading Order Report — Docling",
        f"Source file: {pptx_path}",
        f"Library: docling (DocumentConverter)",
        f"Order semantics: Docling's inferred reading order via layout analysis and",
        f"  reading-order model. This is a DERIVED order, not raw XML element order.",
        f"  Docling groups content into semantic blocks and reorders based on layout.",
        "=" * 80,
    ]

    # Docling organizes content as a tree: body → groups (slides) → texts/pictures/etc.
    # Each node has a self_ref and children with $ref pointers.
    # We build a ref→item map and recursively traverse to get reading order.

    content_map = {}
    for section in ("texts", "tables", "pictures", "key_value_items", "groups"):
        items = data.get(section, [])
        if isinstance(items, list):
            for item in items:
                ref = item.get("self_ref", "")
                if ref:
                    content_map[ref] = item

    # Also map the body itself
    body = data.get("body", {})
    if body.get("self_ref"):
        content_map[body["self_ref"]] = body

    elem_counter = [0]  # mutable counter for element numbering per slide
    current_slide = [None]  # track current slide number

    def traverse_node(ref, indent=0):
        """Recursively traverse docling content tree in reading order."""
        item = content_map.get(ref)
        if item is None:
            lines.append(f"{'  ' * indent}(unresolved ref: {ref})")
            return

        label = item.get("label", "unknown")
        name = item.get("name", "")
        text = item.get("text", "")
        self_ref = item.get("self_ref", ref)

        # Extract provenance (page number and bbox)
        prov = item.get("prov", [])
        page_no = None
        bbox_str = ""
        for p in prov:
            page_no = p.get("page_no")
            bbox = p.get("bbox")
            if bbox:
                coord_origin = bbox.get("coord_origin", "")
                bbox_str = (f"bbox=({bbox.get('l', '?')}, {bbox.get('t', '?')}, "
                            f"{bbox.get('r', '?')}, {bbox.get('b', '?')}) "
                            f"origin={coord_origin}")
            break

        # Emit slide header when page changes
        if page_no and page_no != current_slide[0]:
            current_slide[0] = page_no
            lines.append("")
            lines.append(f"--- Slide {page_no} ---")
            lines.append("")
            elem_counter[0] = 0

        # For group nodes (slides), print header then recurse into children
        is_group = label in ("chapter", "group", "unspecified") and "groups" in self_ref
        prefix = "  " * indent

        if is_group:
            lines.append(f"{prefix}[Group] Label: {label} | Name: {name} | Ref: {self_ref}")
        else:
            lines.append(f"{prefix}[Slide {current_slide[0] or '?'}] Element #{elem_counter[0]} | "
                         f"Label: {label} | Ref: {self_ref}")
            if name:
                lines.append(f"{prefix}  Name: {name}")
            if bbox_str:
                lines.append(f"{prefix}  Position: {bbox_str}")
            if text:
                lines.append(f"{prefix}  Text: {text[:200]}")
            lines.append("")
            elem_counter[0] += 1

        # Recurse into children (this preserves docling's reading order)
        children = item.get("children", [])
        for child in children:
            child_ref = child.get("$ref", "")
            if child_ref:
                traverse_node(child_ref, indent + 1 if is_group else indent)

    # Start traversal from body's children
    for child in body.get("children", []):
        child_ref = child.get("$ref", "")
        if child_ref:
            traverse_node(child_ref)

    text_out = "\n".join(lines)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text_out)

    print(f"[docling] Reading order saved to: {output_path}")
    return output_path


if __name__ == "__main__":
    if "--reading" in sys.argv:
        pptx_files = globmod.glob("*.pptx")
        if not pptx_files:
            print("No PPTX files found in current directory.")
        for pf in pptx_files:
            extract_reading_order(pf)
    else:
        converter = DocumentConverter()
        doc = converter.convert("test-with-images.pptx")
        markdown_text = doc.document.export_to_markdown()

        with open("docling-parser/test-with-images.md", "w", encoding="utf-8") as f:
            f.write(markdown_text)

        data = doc.document.export_to_dict()

        with open("docling-parser/test-with-images.json", "w", encoding="utf8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
