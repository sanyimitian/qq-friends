"""Process-local workaround for changing UIA control types during QQ refresh."""


def install_uia_wrapper_fix():
    from pywinauto.controls.uiawrapper import UIAWrapper, UiaMeta

    if getattr(UiaMeta.find_wrapper, '_qq_single_type_read', False):
        return

    def find_wrapper(element):
        # The bundled implementation reads twice: membership and then lookup.
        # QQ may destroy a node between those reads, changing its type to None.
        control_type = element.control_type
        return UiaMeta.control_type_to_cls.get(control_type, UIAWrapper)

    find_wrapper._qq_single_type_read = True
    UiaMeta.find_wrapper = staticmethod(find_wrapper)
