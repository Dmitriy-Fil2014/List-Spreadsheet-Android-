"""Safe spreadsheet formula evaluator with arithmetic, references and common functions."""
import ast
import operator
import re
import statistics

REF_RE = re.compile(r"(?<![A-Z0-9_])\$?([A-Z]{1,3})\$?(\d+)(?:\s*:\s*\$?([A-Z]{1,3})\$?(\d+))?", re.I)
BINOPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
          ast.Div: operator.truediv, ast.Pow: operator.pow, ast.Mod: operator.mod}
UNARYOPS = {ast.UAdd: operator.pos, ast.USub: operator.neg}
CMPOPS = {ast.Eq: operator.eq, ast.NotEq: operator.ne, ast.Lt: operator.lt,
          ast.LtE: operator.le, ast.Gt: operator.gt, ast.GtE: operator.ge}


class FormulaError(ValueError):
    pass


def col_to_num(col):
    n = 0
    for ch in col.upper():
        n = n * 26 + ord(ch) - 64
    return n


def _num_to_col(n):
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def expand_range(ref):
    a, b = ref.split(":")
    ma, mb = re.fullmatch(r"([A-Z]+)(\d+)", a.upper()), re.fullmatch(r"([A-Z]+)(\d+)", b.upper())
    if not ma or not mb:
        raise FormulaError("Некорректный диапазон")
    c1, r1 = col_to_num(ma.group(1)), int(ma.group(2))
    c2, r2 = col_to_num(mb.group(1)), int(mb.group(2))
    if (max(c1, c2)-min(c1, c2)+1) * (max(r1, r2)-min(r1, r2)+1) > 10000:
        raise FormulaError("Слишком большой диапазон")
    return [f"{_num_to_col(c)}{r}" for r in range(min(r1, r2), max(r1, r2)+1)
            for c in range(min(c1, c2), max(c1, c2)+1)]


class FormulaEngine:
    def __init__(self, get_cell):
        self.get_cell = get_cell

    def evaluate(self, formula):
        expression = str(formula).strip()
        if not expression.startswith("="):
            return expression
        expression = expression[1:].replace(";", ",")
        def replace_ref(match):
            first = f"{match.group(1).upper()}{match.group(2)}"
            if match.group(3):
                last = f"{match.group(3).upper()}{match.group(4)}"
                return f'RANGE("{first}:{last}")'
            return f'CELL("{first}")'
        # Preserve quoted text while replacing cell references.
        parts = re.split(r'(".*?(?<!\\)")', expression)
        for i in range(0, len(parts), 2):
            parts[i] = REF_RE.sub(replace_ref, parts[i])
        expression = "".join(parts)
        try:
            return self._node(ast.parse(expression, mode="eval").body)
        except FormulaError:
            raise
        except Exception as exc:
            raise FormulaError(str(exc)) from exc

    def _node(self, node):
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Name):
            if node.id.upper() == "TRUE": return True
            if node.id.upper() == "FALSE": return False
            raise FormulaError("Неизвестное имя: " + node.id)
        if isinstance(node, ast.BinOp) and type(node.op) in BINOPS:
            return BINOPS[type(node.op)](self._node(node.left), self._node(node.right))
        if isinstance(node, ast.UnaryOp) and type(node.op) in UNARYOPS:
            return UNARYOPS[type(node.op)](self._node(node.operand))
        if isinstance(node, ast.Compare):
            left = self._node(node.left)
            for op, comparator in zip(node.ops, node.comparators):
                right = self._node(comparator)
                if type(op) not in CMPOPS or not CMPOPS[type(op)](left, right): return False
                left = right
            return True
        if isinstance(node, ast.BoolOp):
            values = [bool(self._node(v)) for v in node.values]
            return all(values) if isinstance(node.op, ast.And) else any(values)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            name = node.func.id.upper()
            args = [self._node(arg) for arg in node.args]
            if name == "CELL" and len(args) == 1: return _number(self.get_cell(str(args[0])))
            if name == "RANGE" and len(args) == 1:
                return [_number(self.get_cell(ref)) for ref in expand_range(str(args[0]))]
            flat = []
            for arg in args: flat.extend(arg if isinstance(arg, list) else [arg])
            nums = [float(x) for x in flat if _is_number(x)]
            if name == "SUM": return sum(nums)
            if name == "AVERAGE": return sum(nums)/len(nums) if nums else 0
            if name == "MIN": return min(nums) if nums else 0
            if name == "MAX": return max(nums) if nums else 0
            if name == "COUNT": return len(nums)
            if name == "COUNTA": return sum(x not in ("", None) for x in flat)
            if name == "MEDIAN": return statistics.median(nums) if nums else 0
            if name == "STDEV": return statistics.stdev(nums) if len(nums)>1 else 0
            if name == "VAR": return statistics.variance(nums) if len(nums)>1 else 0
            if name == "IF" and len(args) in (2,3): return args[1] if args[0] else (args[2] if len(args)==3 else False)
            if name == "AND": return all(bool(x) for x in flat)
            if name == "OR": return any(bool(x) for x in flat)
            if name == "NOT" and len(args)==1: return not bool(args[0])
            if name == "ABS" and len(nums)==1: return abs(nums[0])
            if name == "ROUND" and len(args) in (1,2): return round(float(args[0]),int(args[1]) if len(args)==2 else 0)
            if name == "SQRT" and len(nums)==1: return nums[0]**0.5
            if name == "POWER" and len(args)==2: return float(args[0])**float(args[1])
            if name == "LEN" and len(args)==1: return len(str(args[0]))
            if name == "UPPER" and len(args)==1: return str(args[0]).upper()
            if name == "LOWER" and len(args)==1: return str(args[0]).lower()
            if name == "TRIM" and len(args)==1: return " ".join(str(args[0]).split())
            if name == "LEFT" and len(args) in (1,2): return str(args[0])[:int(args[1]) if len(args)==2 else 1]
            if name == "RIGHT" and len(args) in (1,2): return str(args[0])[-(int(args[1]) if len(args)==2 else 1):]
            if name == "MID" and len(args)==3: return str(args[0])[int(args[1])-1:int(args[1])-1+int(args[2])]
            if name in ("CONCATENATE","CONCAT"): return "".join(map(str,flat))
            if name == "COUNTIF" and len(args)==2:
                return sum(1 for x in (args[0] if isinstance(args[0],list) else [args[0]]) if str(x)==str(args[1]))
            raise FormulaError(f"Неизвестная функция или аргументы: {name}")
        raise FormulaError("Недопустимое выражение в формуле")


def _is_number(value):
    try: float(value); return value not in ("", None)
    except (TypeError, ValueError): return False


def _number(value):
    if value in ("", None): return 0
    try: return float(value)
    except (TypeError, ValueError): return value
