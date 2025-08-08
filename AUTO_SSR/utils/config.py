# utils/config.py

import json
from .resolver import resolve_dir

class ConfigDict(dict):
    def __init__(self, filepath=None):
        self._filepath = filepath or ASSETS_DIR / "data.json"

        try:
            with open(self._filepath, "r", encoding="utf-8") as f:
                initial_data = json.load(f)
                super().__init__(initial_data)
                self._wrap_nested()  # Ensure nested dicts are wrapped for dot-access
        except (FileNotFoundError, json.JSONDecodeError):
            super().__init__()

        self.save()  # Ensure file exists and is synced

    def save(self):
        with open(self._filepath, "w", encoding="utf-8") as f:
            json.dump(self, f, indent=4)

    # --- 🔌 Auto-save dict mutation hooks ---
    def __setitem__(self, key, value):
        super().__setitem__(key, value)
        self._wrap_nested()  # Wrap any newly added dicts
        self.save()

    def update(self, *args, **kwargs):
        super().update(*args, **kwargs)
        self._wrap_nested()  # Wrap nested dicts after update
        self.save()

    def pop(self, *args, **kwargs):
        result = super().pop(*args, **kwargs)
        self.save()
        return result

    def clear(self):
        super().clear()
        self.save()

    def reload(self):
        try:
            with open(self._filepath, "r", encoding="utf-8") as f:
                self.clear()
                self.update(json.load(f))
            self._wrap_nested()  # Wrap any newly loaded nested dicts
        except (FileNotFoundError, json.JSONDecodeError):
            pass

    # --- 🧠 Enable dot-access like SimpleNamespace ---
    def __getattr__(self, name):
        if name in self:
            return self[name]
        raise AttributeError(f"'ConfigDict' has no attribute '{name}'")

    def __setattr__(self, name, value):
        if name.startswith('_') or name in self.__dict__:
            super().__setattr__(name, value)
        else:
            self[name] = value  # triggers autosave

    def __delattr__(self, name):
        if name in self:
            del self[name]
            self.save()
        else:
            super().__delattr__(name)

    # --- 🚀 Wrap any nested dicts to support dot-access at all levels ---
    def _wrap_nested(self):
        for k, v in self.items():
            if isinstance(v, dict) and not isinstance(v, ConfigDict):
                self[k] = ConfigDict(v, filepath=self._filepath)  # wrap nested dicts

# Set the path to assets
ASSETS_DIR = resolve_dir("assets")

# Instantiate the auto-loading, auto-saving config
REPORT_DATA = ConfigDict()