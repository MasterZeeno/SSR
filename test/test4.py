import inspect
import re

DEFAULT_FONT_SIZE = 16
CURR_FONT_SIZE = DEFAULT_FONT_SIZE

def flatten(obj, dimen_keys):
    if not obj:
        if dimen_keys:
            return dict.fromkeys(dimen_keys, 0)
        else:
            return None
    
    res = dict.fromkeys(dimen_keys, None)
    def _flatten(o):
        if isinstance(o, dict):
            default = o.get('default', 0)
            for k in dimen_keys:
                res[k] = o.get(k, default)
        elif isinstance(o, (list, tuple, set)):
            for i in o:
                if isinstance(i, str):
                    ps = re.split(r'[^a-z0-9.%+-]+', i, flags=re.I)
                    if len(ps) > 1:
                        _flatten(ps)
                        
                for k in dimen_keys:
                    if res[k] is None:
                        res[k] = i if i else 0
                        break
                        
            for idx, key in enumerate(dimen_keys, -2):
                if res[key] is None:
                    res[key] = res[dimen_keys[0 if idx < 0 else idx]]
    
    _flatten(obj)
    return res

def to_css_val(val):
    s = str(val).strip().lower()
    t = re.search(r'([^0-9.]+)$', s)
    u = t.group(1) if t else 'px'
    n = re.search(r'[-+]?\d*\.\d+|[-+]?\d+', s)
    f = float(n.group()) if n else 0
    if u.endswith('em'):
        u = 'px'
        f *= CURR_FONT_SIZE if u == 'em' else DEFAULT_FONT_SIZE
    return f"{int(f) if f.is_integer() else round(f, 3)}{u}"

def __space_resolver(*args, **kwargs):
    prop = inspect.stack()[1].function.replace('set_', '')
    if prop not in ('padding', 'margin'):
        return
    
    dimen_keys = ('top', 'left', 'bottom', 'right')
        
    return ';'.join(f'{prop}-{key}:{val if val == "auto" else to_css_val(val)}'
        for key, val in flatten(args or kwargs).items())

def set_padding(*args, **kwargs):
    return __space_resolver(*args, **kwargs)

def set_margin(*args, **kwargs):
    return __space_resolver(*args, **kwargs)

# def set_border(*args, **kwargs):
    # dotted - Defines a dotted border
    # dashed - Defines a dashed border
    # solid - Defines a solid border
    # double - Defines a double border
    # groove - Defines a 3D grooved border. The effect depends on the border-color value
    # ridge - Defines a 3D ridged border. The effect depends on the border-color value
    # inset - Defines a 3D inset border. The effect depends on the border-color value
    # outset - Defines a 3D outset border. The effect depends on the border-color value
    # none - Defines no border
    # hidden - Defines a hidden border
    
    
def tester(*args, **kwargs):
    # top, left, bottom, right = ([*args] + [None] * 4)[:4]
    # data = {"top": 0, "left": 0, "bottom": 0, "right": 0}
    v = args if args else [0]
    data = dict(zip(('top', 'left', 'bottom', 'right'), (v + (v[1::2] if len(v) > 2 else v) * (4 // len(v)) if len(v) < 4 else [v[1]])[:4]))
    data = kwargs
    # for i, k in enumerate(data.keys()):
        # if i < len(args):
            # data[k] = args[i]
    print(data)
    # [top, left, bottom, right] = kwargs.items()
    # print(top, left, bottom, right)

tester(top=34, right=5, bottom=4, left=2)
tester(3, 4, 7)
tester(3, 4)
tester(3)
tester()
# tester(top=34, right=5, bottom=4, left=2)
# tester()

    