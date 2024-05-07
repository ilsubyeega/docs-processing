from __future__ import annotations
import uno
# noinspection PyUnresolvedReferences
from msgbox import MsgBox  # this is work as expected

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from com.sun.star.script.provider import XScriptContext
    from com.sun.star.lang import XComponent
    from com.sun.star.drawing import DrawPages, DrawPage, DrawingDocument, Shape

    # noinspection SpellCheckingInspection
    XSCRIPTCONTEXT = XScriptContext()
    ThisComponent = XComponent()


def remove_same_shapes_for_every_page():
    def get_shape_properties(shape: Shape) -> dict:
        return {
            "name": f'{shape.Name} ({shape.getName()}',
            "width": shape.getSize().Width.real,
            "height": shape.getSize().Height.real,
            "type": shape.getShapeType(),
            "x": shape.getPosition().X,
            "y": shape.getPosition().Y
        }

    doc: DrawingDocument = XSCRIPTCONTEXT.getDocument()
    pages: DrawPages = doc.getDrawPages()

    # save all shapes with page key
    shapes_with_pages_key = []
    for page in [pages.getByIndex(i) for i in range(pages.getCount())]:
        for i in range(page.getCount()):
            shape = page.getByIndex(i)
            shapes_with_pages_key.append((shape, page, get_shape_properties(shape)))

    # get selected shapes
    to_remove = []
    selection = doc.getCurrentSelection()
    if selection:
        for i in range(selection.getCount()):
            shape: Shape = selection.getByIndex(i)

            # check duplicates
            this_props = get_shape_properties(shape)
            duplicates = []
            removed_shapes = []
            for s, p, props in shapes_with_pages_key:
                if s in removed_shapes:
                    continue  # already marked for removal.

                weak_check = this_props["width"] == props["width"] \
                             and this_props["height"] == props["height"] \
                             and this_props["type"] == props["type"]

                pos_check = this_props["x"] == props["x"] \
                            and this_props["y"] == props["y"]

                if weak_check:
                    duplicates.append((s, p, props, weak_check, pos_check))
                    removed_shapes.append(s)

            if len(duplicates) > 1:
                for s, p, props, weak_check, pos_check in duplicates:
                    shapes_with_pages_key.remove((s, p, props))
                    to_remove.append((s, p, props, weak_check, pos_check))

    # remove duplicates
    text = f"Removed {len(to_remove)} duplicates"
    for s, p, props, weak_check, pos_check in to_remove:
        p.remove(s)
        text += f"\n{props["name"]} - {props["type"]} - Size {props["width"]}x{props["height"]} - Pos {props["x"]}, {props["y"]}"

    msgbox = MsgBox(XSCRIPTCONTEXT.getComponentContext())
    msgbox.addButton("aight")
    msgbox.renderFromBoxSize(size=300)
    msgbox.numberOflines = 6
    msgbox.show(text, 0, "Dupe Removal")
