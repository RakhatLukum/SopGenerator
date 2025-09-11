import os
from typing import Any, Optional

try:
    import streamlit as st  # type: ignore
except Exception:
    st = None  # Streamlit may not be available in some contexts

_ENV_MAP = {
    "api_key": ["LLM_API_KEY", "GOVPLAN_API_KEY", "API_KEY", "OPENROUTER_API_KEY"],
}


def get(key: str, default: Optional[Any] = None) -> Optional[Any]:
    # Prefer Streamlit secrets if available
    if st is not None:
        try:
            if key in st.secrets:
                return st.secrets[key]
        except Exception:
            pass
        for env_key in _ENV_MAP.get(key, []):
            try:
                if env_key in st.secrets:
                    return st.secrets[env_key]
            except Exception:
                pass

    # Fallback to environment variables via mapping
    for env_key in _ENV_MAP.get(key, []):
        value = os.getenv(env_key)
        if value:
            return value

    return default 