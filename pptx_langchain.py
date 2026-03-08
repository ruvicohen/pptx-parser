import os
import sys
import glob as globmod

from langchain_community.document_loaders import UnstructuredPowerPointLoader


def extract_reading_order(pptx_path, output_dir="reading_order_outputs"):
    """Extract and save the reading order of elements using the Unstructured library.

    Reading order for Unstructured: The unstructured library's partition_pptx
    function returns elements in the order it parses them from the PPTX.
    This generally follows slide order and shape iteration order within each slide,
    but unstructured may reorder elements internally (e.g., titles before body text).
    This is the library's OWN parsing order — not raw XML order.
    Unstructured exposes element category (Title, Text, ListItem, etc.),
    page number, and parent_id for hierarchy, but does NOT expose shape-level
    position coordinates or bounding boxes.
    """
    os.makedirs(output_dir, exist_ok=True)

    # Use unstructured directly (more reliable than the langchain wrapper)
    from unstructured.partition.pptx import partition_pptx
    elements = partition_pptx(pptx_path)

    pptx_name = os.path.splitext(os.path.basename(pptx_path))[0]
    output_filename = f"langchain_unstructured__{pptx_name}.txt"
    output_path = os.path.join(output_dir, output_filename)

    lines = [
        f"Reading Order Report — Unstructured",
        f"Source file: {pptx_path}",
        f"Library: unstructured (partition_pptx)",
        f"Order semantics: Elements in the order unstructured's partition_pptx returns them.",
        f"  This is the library's own parsing order. Unstructured may reorder elements",
        f"  (e.g., promoting titles). No position/bbox data is exposed by this library.",
        "=" * 80,
    ]

    current_page = None
    for elem_index, elem in enumerate(elements):
        meta = elem.metadata
        page = meta.page_number

        if page != current_page:
            current_page = page
            lines.append("")
            lines.append(f"--- Slide {current_page or '?'} ---")
            lines.append("")

        elem_type = type(elem).__name__
        category_depth = meta.category_depth
        parent_id = meta.parent_id
        text = str(elem).strip()

        lines.append(f"  [Slide {current_page or '?'}] Element #{elem_index} | "
                     f"Type: {elem_type} | Depth: {category_depth}")
        if parent_id:
            lines.append(f"    Parent ID: {parent_id}")
        if text:
            lines.append(f"    Text: {text[:200]}")
        lines.append("")

    text_out = "\n".join(lines)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text_out)

    print(f"[unstructured] Reading order saved to: {output_path}")
    return output_path


if __name__ == "__main__":
    if "--reading" in sys.argv:
        pptx_files = globmod.glob("*.pptx")
        if not pptx_files:
            print("No PPTX files found in current directory.")
        for pf in pptx_files:
            extract_reading_order(pf)
    else:
        # Original behavior
        loader = UnstructuredPowerPointLoader("test-with-images.pptx", mode="elements", strategy="fast")
        documents = loader.load()
        text = ""
        for d in documents:
            print(d.metadata)
            print(d.page_content)
            print("------")
            text += d.page_content + '\n'

        with open("langchain-parser/test-with-images.txt", "w", encoding="utf-8") as f:
            f.write(text)