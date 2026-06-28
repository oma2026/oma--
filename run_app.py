# -*- coding: utf-8 -*-
import streamlit as st
import streamlit_title_override
from streamlit.web import cli as stcli

streamlit_title_override.apply(st)

if __name__ == "__main__":
    stcli.main_run(
        "app.py",
        args=["--server.address", "0.0.0.0", "--server.port", "8501"],
        flag_options={},
    )
