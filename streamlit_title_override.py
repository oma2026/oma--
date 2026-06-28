# -*- coding: utf-8 -*-
OLD_TITLE = "二手車進口報價系統"
NEW_TITLE = "進口報價系統二手車"

def apply(st):
    original_set_page_config = st.set_page_config
    def set_page_config(*args, **kwargs):
        if kwargs.get("page_title") == OLD_TITLE:
            kwargs["page_title"] = NEW_TITLE
        return original_set_page_config(*args, **kwargs)
    st.set_page_config = set_page_config

    original_sidebar_title = st.sidebar.title
    def sidebar_title(body=None, *args, **kwargs):
        if body == OLD_TITLE:
            body = NEW_TITLE
        return original_sidebar_title(body, *args, **kwargs)
    st.sidebar.title = sidebar_title
