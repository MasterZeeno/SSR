import os

def set_env(var, value=None):
    if isinstance(var, dict):
        for _var, _val in var.items():
            update(_var, _val)
    elif isinstance(var, str):
        os.environ[var.upper().strip()] = (
            str(value) if value is not None else ''
        )
    else:
        raise TypeError(f"Unsupported type: {type(var)}")
        