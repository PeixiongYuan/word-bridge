#!/usr/bin/env python3
"""Turn LaTeX math into UnicodeMath, the linear format Word builds its native equations from.

Word's equation engine reads a line such as  〖(a+b)/(c)〗  and builds a real equation object out of it.
It does not read LaTeX from a script, so latex_to_unicodemath() translates a defined subset of LaTeX.
Anything outside that subset raises MathError instead of being guessed at.

Standard library only.
"""
import re


class MathError(ValueError):
    """The formula cannot be converted; the message says what to change."""


# Characters with a meaning of their own in UnicodeMath
GROUP_OPEN, GROUP_CLOSE = "〖", "〗"      # 〖 〗 invisible grouping
APPLY = "⁡"                                    # function application
OPERAND = "▒"                                  # ▒ glues an n-ary operator to its operand
BELOW, ABOVE = "┬", "┴"                   # ┬ ┴ limits below and above
OPEN_ANY, CLOSE_ANY = "├", "┤"            # ├ ┤ introduce an arbitrary delimiter
MATRIX, ARRAY, ATOP = "■", "█", "¦"  # ■ █ ¦
EMPTY = "​"                                    # zero-width space, stands in for an empty argument
# Linear-format characters that must not survive in a built equation
LEFTOVER_CHARS = GROUP_OPEN + GROUP_CLOSE + APPLY + OPERAND + BELOW + ABOVE + OPEN_ANY + CLOSE_ANY + MATRIX + ARRAY + ATOP

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ϵ", "varepsilon": "ε", "zeta": "ζ", "eta": "η",
    "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ", "varkappa": "ϰ", "lambda": "λ", "mu": "μ", "nu": "ν",
    "xi": "ξ", "omicron": "ο", "pi": "π", "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ", "sigma": "σ", "varsigma": "ς",
    "tau": "τ", "upsilon": "υ", "phi": "ϕ", "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ", "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ",
    "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}

# name -> (character, class). "bin" (signs), "rel" and "punct" end the operand of a sum or a function.
SYMBOLS = {
    "pm": ("±", "bin"), "mp": ("∓", "bin"),
    "times": ("×", "ord"), "cdot": ("⋅", "ord"), "div": ("÷", "ord"), "ast": ("∗", "ord"), "star": ("⋆", "ord"),
    "circ": ("∘", "ord"), "bullet": ("∙", "ord"), "otimes": ("⊗", "ord"), "oplus": ("⊕", "ord"), "odot": ("⊙", "ord"),
    "cup": ("∪", "ord"), "cap": ("∩", "ord"), "setminus": ("∖", "ord"), "wedge": ("∧", "ord"), "land": ("∧", "ord"),
    "vee": ("∨", "ord"), "lor": ("∨", "ord"),
    "leq": ("≤", "rel"), "le": ("≤", "rel"), "geq": ("≥", "rel"), "ge": ("≥", "rel"), "neq": ("≠", "rel"), "ne": ("≠", "rel"),
    "approx": ("≈", "rel"), "sim": ("∼", "rel"), "simeq": ("≃", "rel"), "cong": ("≅", "rel"), "equiv": ("≡", "rel"),
    "propto": ("∝", "rel"), "ll": ("≪", "rel"), "gg": ("≫", "rel"), "doteq": ("≐", "rel"),
    "in": ("∈", "rel"), "notin": ("∉", "rel"), "ni": ("∋", "rel"), "subset": ("⊂", "rel"), "subseteq": ("⊆", "rel"),
    "supset": ("⊃", "rel"), "supseteq": ("⊇", "rel"), "perp": ("⊥", "rel"), "parallel": ("∥", "rel"),
    "to": ("→", "rel"), "rightarrow": ("→", "rel"), "leftarrow": ("←", "rel"), "gets": ("←", "rel"),
    "leftrightarrow": ("↔", "rel"), "Rightarrow": ("⇒", "rel"), "Leftarrow": ("⇐", "rel"), "Leftrightarrow": ("⇔", "rel"),
    "implies": ("⇒", "rel"), "iff": ("⇔", "rel"), "mapsto": ("↦", "rel"), "longrightarrow": ("⟶", "rel"),
    "uparrow": ("↑", "ord"), "downarrow": ("↓", "ord"),
    "mid": ("∣", "rel"), "vdash": ("⊢", "rel"), "models": ("⊨", "rel"),
    "infty": ("∞", "ord"), "partial": ("∂", "ord"), "nabla": ("∇", "ord"), "forall": ("∀", "ord"), "exists": ("∃", "ord"),
    "nexists": ("∄", "ord"), "emptyset": ("∅", "ord"), "varnothing": ("∅", "ord"), "neg": ("¬", "ord"), "lnot": ("¬", "ord"),
    "ell": ("ℓ", "ord"), "hbar": ("ℏ", "ord"), "Re": ("ℜ", "ord"), "Im": ("ℑ", "ord"), "aleph": ("ℵ", "ord"),
    "imath": ("ı", "ord"), "jmath": ("ȷ", "ord"), "angle": ("∠", "ord"), "triangle": ("△", "ord"),
    "prime": ("′", "ord"), "degree": ("°", "ord"), "dagger": ("†", "ord"), "ddagger": ("‡", "ord"),
    "ldots": ("…", "ord"), "dots": ("…", "ord"), "dotsc": ("…", "ord"), "cdots": ("⋯", "ord"), "dotsb": ("⋯", "ord"),
    "vdots": ("⋮", "ord"), "ddots": ("⋱", "ord"),
    "vert": ("|", "ord"), "lvert": ("|", "ord"), "rvert": ("|", "ord"),
    "Vert": ("‖", "ord"), "lVert": ("‖", "ord"), "rVert": ("‖", "ord"), "|": ("‖", "ord"),
    "langle": ("⟨", "open"), "rangle": ("⟩", "close"), "lfloor": ("⌊", "open"), "rfloor": ("⌋", "close"),
    "lceil": ("⌈", "open"), "rceil": ("⌉", "close"), "lbrace": ("{", "open"), "rbrace": ("}", "close"),
    "{": ("{", "open"), "}": ("}", "close"), "lbrack": ("[", "open"), "rbrack": ("]", "close"),
    "backslash": ("∖", "ord"), "top": ("⊤", "ord"), "bot": ("⊥", "ord"), "intercal": ("⊺", "ord"),
    "triangleq": ("≜", "rel"), "coloneqq": ("≔", "rel"), "lesssim": ("≲", "rel"), "gtrsim": ("≳", "rel"),
    "nmid": ("∤", "rel"), "prec": ("≺", "rel"), "succ": ("≻", "rel"), "preceq": ("⪯", "rel"), "succeq": ("⪰", "rel"),
    "therefore": ("∴", "ord"), "because": ("∵", "ord"), "sqcup": ("⊔", "ord"), "ominus": ("⊖", "ord"), "oslash": ("⊘", "ord"),
    "checkmark": ("✓", "ord"), "square": ("□", "ord"), "diamond": ("⋄", "ord"), "hslash": ("ℏ", "ord"), "wp": ("℘", "ord"),
    "%": ("%", "ord"), "$": ("$", "ord"), "&": ("\\&", "ord"), "#": ("\\#", "ord"), "_": ("\\_", "ord"),
}

NARY = {
    "sum": "∑", "prod": "∏", "coprod": "∐", "int": "∫", "iint": "∬", "iiint": "∭", "oint": "∮",
    "bigcup": "⋃", "bigcap": "⋂", "bigvee": "⋁", "bigwedge": "⋀", "bigoplus": "⨁", "bigotimes": "⨂",
}

# Function names; True means that a subscript goes underneath, as with lim
FUNCTIONS = {name: False for name in (
    "sin cos tan cot sec csc arcsin arccos arctan sinh cosh tanh coth log ln lg exp det dim ker deg gcd hom arg Pr").split()}
FUNCTIONS.update({name: True for name in "lim max min sup inf limsup liminf argmax argmin".split()})
FUNCTION_LABELS = {"limsup": "lim sup", "liminf": "lim inf", "argmax": "arg max", "argmin": "arg min"}

ACCENTS = {
    "hat": "̂", "widehat": "̂", "tilde": "̃", "widetilde": "̃", "bar": "̅", "dot": "̇",
    "ddot": "̈", "vec": "⃗", "check": "̌", "breve": "̆", "acute": "́", "grave": "̀",
}

SPACES = {
    ",": " ", "thinspace": " ", ":": " ", "medspace": " ", ";": " ", "thickspace": " ",
    " ": " ", "quad": " ", "qquad": "  ", "enspace": " ", "!": "", "negthinspace": "",
}

IGNORED = set((
    "displaystyle textstyle scriptstyle scriptscriptstyle limits nolimits nonumber notag allowbreak "
    "big Big bigg Bigg bigl bigr Bigl Bigr biggl biggr Biggl Biggr bigm Bigm biggm Biggm hline").split())

TEXT_COMMANDS = ("text", "textrm", "textup", "textnormal", "textit", "textbf", "textsf", "mbox", "mathrm", "rm")
PASS_THROUGH = ("mathit", "mathnormal", "mathord", "mathbin", "mathrel", "mathpunct", "mathopen", "mathclose", "mathinner")
WRAPPERS = ("equation", "equation*", "displaymath", "math")
MATRIX_ENVS = {
    "matrix": ("", ""), "smallmatrix": ("", ""), "array": ("", ""), "pmatrix": ("(", ")"), "bmatrix": ("[", "]"),
    "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"), "Vmatrix": ("‖", "‖"),
}
ARRAY_ENVS = ("aligned", "align", "align*", "split", "gathered", "gather", "gather*", "eqnarray", "eqnarray*", "alignedat")
IMPLICIT_PAIRS = {"(": ")", "[": "]", "{": "}", "⟨": "⟩", "⌊": "⌋", "⌈": "⌉"}
BARS = ("|", "‖")
NOT = {"=": "≠", "<": "≮", ">": "≯", "in": "∉", "equiv": "≢", "subset": "⊄", "subseteq": "⊈", "sim": "≁", "approx": "≉",
       "mid": "∤", "parallel": "∦", "exists": "∄", "leq": "≰", "geq": "≱", "le": "≰", "ge": "≱"}
DELIMITER_NAMES = {
    "(": "(", ")": ")", "[": "[", "]": "]", "|": "|", "/": "/", ".": "",
    "\\{": "{", "\\}": "}", "\\lbrace": "{", "\\rbrace": "}", "\\lbrack": "[", "\\rbrack": "]",
    "\\|": "‖", "\\Vert": "‖", "\\lVert": "‖", "\\rVert": "‖", "\\vert": "|", "\\lvert": "|", "\\rvert": "|",
    "\\langle": "⟨", "\\rangle": "⟩", "<": "⟨", ">": "⟩", "\\lfloor": "⌊", "\\rfloor": "⌋", "\\lceil": "⌈", "\\rceil": "⌉",
}


def _alphabet(upper, lower=None, digit=None, exceptions=None):
    table = {}
    for i in range(26):
        table[chr(65 + i)] = chr(upper + i)
        if lower is not None:
            table[chr(97 + i)] = chr(lower + i)
    if digit is not None:
        for i in range(10):
            table[chr(48 + i)] = chr(digit + i)
    table.update(exceptions or {})
    return table


def _greek_alphabet(upper, lower):
    table = {}
    for i, ch in enumerate("ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡϴΣΤΥΦΧΨΩ"):
        table[ch] = chr(upper + i)
    for i, ch in enumerate("αβγδεζηθικλμνξοπρςστυφχψω"):
        table[ch] = chr(lower + i)
    for ch, offset in (("∂", 25), ("ϵ", 26), ("ϑ", 27), ("ϰ", 28), ("ϕ", 29), ("ϱ", 30), ("ϖ", 31)):
        table[ch] = chr(lower + offset)
    table["∇"] = chr(upper + 25)
    return table


STYLES = {
    "mathbf": {**_alphabet(0x1D400, 0x1D41A, 0x1D7CE), **_greek_alphabet(0x1D6A8, 0x1D6C2)},
    "boldsymbol": {**_alphabet(0x1D468, 0x1D482, 0x1D7CE), **_greek_alphabet(0x1D71C, 0x1D736)},
    "mathbb": _alphabet(0x1D538, 0x1D552, 0x1D7D8, {"C": "ℂ", "H": "ℍ", "N": "ℕ", "P": "ℙ", "Q": "ℚ", "R": "ℝ", "Z": "ℤ"}),
    "mathcal": _alphabet(0x1D49C, 0x1D4B6, None, {"B": "ℬ", "E": "ℰ", "F": "ℱ", "H": "ℋ", "I": "ℐ", "L": "ℒ", "M": "ℳ", "R": "ℛ",
                                                    "e": "ℯ", "g": "ℊ", "o": "ℴ"}),
    "mathfrak": _alphabet(0x1D504, 0x1D51E, None, {"C": "ℭ", "H": "ℌ", "I": "ℑ", "R": "ℜ", "Z": "ℨ"}),
    "mathsf": _alphabet(0x1D5A0, 0x1D5BA, 0x1D7E2),
    "mathtt": _alphabet(0x1D670, 0x1D68A, 0x1D7F6),
}
STYLES["bm"] = STYLES["boldsymbol"]
STYLES["mathscr"] = STYLES["mathcal"]
STYLES["bf"] = STYLES["mathbf"]

_COMMAND = re.compile(r"\\([A-Za-z]+|.)", re.S)
_NUMBER = re.compile(r"\d+(?:\.\d+)?|\.\d+")
_UNSUPPORTED = ("{} is not supported by the LaTeX converter. Rewrite the formula without it, "
                "or pass it in Word's linear format (UnicodeMath) with format=\"unicodemath\".")


class _Parser:
    def __init__(self, src):
        self.s = src
        self.i = 0

    # ------------------------------------------------------------ reading the source

    def skip_ws(self):
        while self.i < len(self.s):
            ch = self.s[self.i]
            if ch.isspace():
                self.i += 1
            elif ch == "%":
                end = self.s.find("\n", self.i)
                self.i = len(self.s) if end < 0 else end + 1
            else:
                break

    def at_end(self):
        self.skip_ws()
        return self.i >= len(self.s)

    def peek(self):
        self.skip_ws()
        return self.s[self.i] if self.i < len(self.s) else ""

    def peek_command(self):
        """Name of the command at the cursor, or None."""
        if self.peek() != "\\":
            return None
        m = _COMMAND.match(self.s, self.i)
        return m.group(1) if m else ""

    def take_command(self):
        m = _COMMAND.match(self.s, self.i)
        if not m:
            raise MathError("The formula ends with a lone backslash.")
        self.i = m.end()
        return m.group(1)

    def expect(self, ch, what):
        if self.peek() != ch:
            raise MathError(f"Expected {ch} {what}.")
        self.i += 1

    def raw_group(self, what):
        """The text of a {...} argument, taken literally (for \\text and friends)."""
        self.expect("{", f"after {what}")
        out, depth = [], 1
        while self.i < len(self.s):
            ch = self.s[self.i]
            self.i += 1
            if ch == "\\" and self.i < len(self.s):
                nxt = self.s[self.i]
                if nxt in "{}%&#_$ ":
                    out.append(nxt)
                    self.i += 1
                    continue
                if nxt in ",;:!":          # spacing commands, as in arg\,max
                    out.append("" if nxt == "!" else " ")
                    self.i += 1
                    continue
                raise MathError(f"A command inside {what}{{...}} is not supported; keep only plain text there.")
            if ch == "$":
                raise MathError(f"Math inside {what}{{...}} is not supported; end the text and write the math outside it.")
            if ch == "{":
                depth += 1
                continue
            if ch == "}":
                depth -= 1
                if depth == 0:
                    return "".join(out).replace("~", " ")
                continue
            out.append(ch)
        raise MathError(f"Missing }} after {what}{{.")

    def environment_name(self):
        return self.raw_group("\\begin or \\end").strip()

    def argument(self, what):
        """One required argument: a {...} group, or a single token as in \\frac12 and x^2."""
        if self.at_end():
            raise MathError(f"Missing argument {what}.")
        if self.peek() == "{":
            self.i += 1
            nodes = self.sequence()
            self.expect("}", f"to close the argument {what}")
            return nodes
        if self.peek() in "}&^_":
            raise MathError(f"Missing argument {what}.")
        if self.peek().isdigit():       # \frac12 is one over two
            self.i += 1
            return [("num", self.s[self.i - 1])]
        node = self.atom()
        return [] if node is None else [node]

    def delimiter(self, what):
        self.skip_ws()
        m = _COMMAND.match(self.s, self.i) if self.peek() == "\\" else None
        token = m.group(0) if m else self.peek()
        if token not in DELIMITER_NAMES:
            raise MathError(f"{token or 'Nothing'} is not a delimiter that can follow {what}.")
        self.i += len(token)
        return DELIMITER_NAMES[token]

    # ------------------------------------------------------------ grammar

    def rows(self):
        """Cells separated by & and rows separated by \\\\, as inside an environment."""
        rows, row = [], []
        while True:
            row.append(self.sequence())
            if self.peek() == "&":
                self.i += 1
                continue
            if self.peek_command() == "\\":
                self.take_command()
                if self.peek() == "[":       # \\[2pt]
                    end = self.s.find("]", self.i)
                    if end < 0:
                        raise MathError("Missing ] after \\\\[.")
                    self.i = end + 1
                rows.append(row)
                row = []
                continue
            rows.append(row)
            break
        if len(rows) > 1 and rows[-1] == [[]]:
            rows.pop()
        return rows

    def sequence(self):
        """Atoms up to the end of the current group, cell or row."""
        nodes, opens = [], []
        while not self.at_end():
            ch = self.peek()
            if ch in "}&":
                break
            if ch == "\\" and self.peek_command() in ("\\", "right", "end"):
                break
            node = self.atom()
            if node is None:
                continue
            # Fold a matching pair of plain brackets or bars into one unit, so that f(a+b) and |x| stay together
            opening = node[0] == "sym" and node[2] == "open"
            if node[0] == "sym" and node[2] == "close":
                at = len(opens) - 1
                while at >= 0 and opens[at][1] in BARS:      # a bar left open inside brackets was a separator, as in P(A|B)
                    at -= 1
                if at >= 0 and IMPLICIT_PAIRS[opens[at][1]] == node[1]:
                    start, opener = opens[at]
                    del opens[at:]
                    body = nodes[start + 1:]
                    del nodes[start:]
                    node = ("delim", opener, node[1], body)
            elif node[0] == "sym" and node[1] in BARS:
                if opens and opens[-1][1] == node[1]:
                    start, opener = opens.pop()
                    body = nodes[start + 1:]
                    del nodes[start:]
                    node = ("delim", opener, opener, body)
                else:
                    opening = True
            if not opening:
                node = self.scripts(node)
            nodes.append(node)
            if opening:
                opens.append((len(nodes) - 1, node[1]))
        return nodes

    def scripts(self, node):
        sub = sup = None
        primes = 0
        while True:
            ch = self.peek()
            if ch == "\\" and self.peek_command() in ("limits", "nolimits"):
                self.take_command()
            elif ch == "'":
                self.i += 1
                primes += 1
            elif ch == "^":
                self.i += 1
                if sup is not None:
                    raise MathError("Double superscript: use braces to say what the second ^ applies to.")
                sup = self.argument("after ^")
            elif ch == "_":
                self.i += 1
                if sub is not None:
                    raise MathError("Double subscript: use braces to say what the second _ applies to.")
                sub = self.argument("after _")
            else:
                break
        if primes:
            sup = [("sym", "′" * primes, "ord")] + (sup or [])
        if sub is not None and not sub:
            sub = None
        if sup is not None and not sup:
            sup = None
        if sub is None and sup is None:
            return node
        if node[0] in ("nary", "func"):
            return node[:2] + (sub, sup) + node[4:]
        if node[0] == "brace":
            return ("under", [node], sub) if node[1] == "⏟" and sub else ("over", [node], sup or sub)
        return ("script", node, sub, sup)

    def atom(self):
        ch = self.peek()
        if ch == "{":
            self.i += 1
            nodes = self.sequence()
            self.expect("}", "to close a group")
            return ("group", nodes)
        if ch == "\\":
            return self.command()
        m = _NUMBER.match(self.s, self.i)
        if m:
            self.i = m.end()
            return ("num", m.group(0))
        self.i += 1
        if ch in "^_":
            raise MathError(f"{ch} has nothing in front of it to attach to.")
        if ch == "'":
            return ("sym", "′", "ord")
        if ch == "~":
            return ("space", " ")
        if ch == "$":
            raise MathError("Remove the $ signs inside the formula; pass one formula without math delimiters.")
        if ch == "#":
            raise MathError("# is not valid in a formula; write \\# for a literal one.")
        if ch in "+-":
            return ("sym", ch, "bin")
        if ch in "=<>:":
            return ("sym", ch, "rel")
        if ch in ",;":
            return ("sym", ch, "punct")
        if ch in "([":
            return ("sym", ch, "open")
        if ch in ")]":
            return ("sym", ch, "close")
        if ch in '/"@':
            return ("sym", "\\" + ch, "ord")
        return ("sym", ch, "ord")

    def command(self):
        name = self.take_command()
        if name in GREEK:
            return ("sym", GREEK[name], "ord")
        if name in SYMBOLS:
            return ("sym",) + SYMBOLS[name]
        if name in SPACES:
            return ("space", SPACES[name]) if SPACES[name] else None
        if name in IGNORED:
            return None
        if name in ("frac", "dfrac", "tfrac", "cfrac"):
            return ("frac", self.argument(f"of \\{name}"), self.argument(f"of \\{name}"))
        if name in ("binom", "dbinom", "tbinom"):
            return ("binom", self.argument(f"of \\{name}"), self.argument(f"of \\{name}"))
        if name == "sqrt":
            degree = None
            if self.peek() == "[":
                end = self.s.find("]", self.i)
                if end < 0:
                    raise MathError("Missing ] after \\sqrt[.")
                degree = _Parser(self.s[self.i + 1:end]).sequence()
                self.i = end + 1
            return ("sqrt", degree, self.argument("of \\sqrt"))
        if name in NARY:
            return ("nary", NARY[name], None, None)
        if name in FUNCTIONS:
            return ("func", FUNCTION_LABELS.get(name, name), None, None, FUNCTIONS[name])
        if name in ("operatorname", "mathop"):
            below = name == "mathop"
            if name == "operatorname" and self.i < len(self.s) and self.s[self.i] == "*":
                self.i += 1
                below = True
            label = self.raw_group("\\" + name).strip()
            if not label:
                raise MathError(f"\\{name}{{}} needs a name.")
            return ("func", label, None, None, below)
        if name in ACCENTS:
            return ("accent", ACCENTS[name], self.argument(f"of \\{name}"))
        if name in ("overline", "underline"):
            return ("bar", "¯" if name == "overline" else "▁", self.argument(f"of \\{name}"))
        if name in ("overset", "stackrel"):
            top = self.argument(f"of \\{name}")
            return ("over", self.argument(f"of \\{name}"), top)
        if name == "underset":
            bottom = self.argument("of \\underset")
            return ("under", self.argument("of \\underset"), bottom)
        if name in ("underbrace", "overbrace"):
            return ("brace", "⏟" if name == "underbrace" else "⏞", self.argument(f"of \\{name}"))
        if name in TEXT_COMMANDS:
            text = self.raw_group("\\" + name)
            if name in ("mathrm", "rm"):
                text = "".join(text.split())        # math mode ignores spaces
            return ("text", text)
        if name in STYLES:
            return ("styled", name, self.argument(f"of \\{name}"))
        if name in PASS_THROUGH:
            return ("group", self.argument(f"of \\{name}"))
        if name == "left":
            opener = self.delimiter("\\left")
            body = self.sequence()
            if self.peek_command() != "right":
                raise MathError("\\left has no matching \\right.")
            self.take_command()
            return ("delim", opener, self.delimiter("\\right"), body)
        if name == "right":
            raise MathError("\\right has no matching \\left.")
        if name == "begin":
            return self.environment()
        if name == "end":
            raise MathError("\\end has no matching \\begin.")
        if name == "substack":
            self.expect("{", "after \\substack")
            rows = self.rows()
            self.expect("}", "to close \\substack")
            return ("matrix", "aligned", rows)
        if name == "not":
            if self.peek() == "\\":
                target = self.take_command()
            else:
                target = self.peek()
                self.i += 1
            if target not in NOT:
                raise MathError(_UNSUPPORTED.format("\\not before this symbol"))
            return ("sym", NOT[target], "rel")
        if name == "pmod":
            return ("delim", "(", ")", [("text", "mod ")] + self.argument("of \\pmod"))
        if name in ("bmod", "mod"):
            return ("text", " mod ")
        if name == "label":
            self.raw_group("\\label")
            return None
        if name == "tag":
            raise MathError("\\tag is not supported: equation numbers are not added. Leave it out and, if a number is needed, tell the user.")
        raise MathError(_UNSUPPORTED.format("\\" + name))

    def environment(self):
        env = self.environment_name()
        if env in ("array", "alignedat"):
            self.raw_group(f"\\begin{{{env}}}")      # the column specification
        if env in WRAPPERS:
            kind = "wrapper"
        elif env in MATRIX_ENVS or env in ARRAY_ENVS or env == "cases":
            kind = env
        else:
            raise MathError(_UNSUPPORTED.format(f"The environment {env}"))
        rows = self.rows()
        if self.peek_command() != "end":
            raise MathError(f"\\begin{{{env}}} has no matching \\end.")
        self.take_command()
        closing = self.environment_name()
        if closing != env:
            raise MathError(f"\\begin{{{env}}} is closed by \\end{{{closing}}}.")
        if kind == "wrapper":
            return ("group", [("matrix", "aligned", rows)] if len(rows) > 1 or len(rows[0]) > 1 else rows[0][0])
        return ("matrix", kind, rows)


# ---------------------------------------------------------------- writing UnicodeMath

def _group(text):
    return GROUP_OPEN + text + GROUP_CLOSE


def _is_terminator(node):
    return node[0] == "sym" and node[2] in ("bin", "rel", "punct")


def _emit_sequence(nodes):
    out, i = [], 0
    while i < len(nodes):
        node = nodes[i]
        if node[0] in ("nary", "func"):
            # The operand runs up to the next sign, relation or comma at this level
            j = i + 1
            while j < len(nodes) and not _is_terminator(nodes[j]):
                j += 1
            while j > i + 1 and nodes[j - 1][0] == "space":
                j -= 1
            operand = _emit_sequence(nodes[i + 1:j])
            out.append(_emit_nary(node, operand) if node[0] == "nary" else _emit_function(node, operand))
            i = j
        else:
            out.append(_emit(node))
            i += 1
    return "".join(out)


def _script_part(mark, nodes):
    return "" if nodes is None else f"{mark}({_emit_sequence(nodes)})"


def _emit_nary(node, operand):
    _, symbol, sub, sup = node
    return _group(symbol + _script_part("_", sub) + _script_part("^", sup) + OPERAND + _group(operand or EMPTY))


def _quoted(text):
    return '"' + text.replace('"', "”") + '"' if text else ""


def _emit_function(node, operand):
    _, label, sub, sup, below = node
    head = _quoted(label) + _script_part(BELOW if below else "_", sub) + _script_part("^", sup)
    if not operand:
        return _group(head) if (sub or sup) else head
    return _group(head + APPLY + _group(operand))


def _emit(node):
    kind = node[0]
    if kind in ("sym", "num", "space"):
        return node[1]
    if kind == "group":
        inner = _emit_sequence(node[1])
        return inner if len(node[1]) <= 1 else _group(inner)
    if kind == "text":
        return _quoted(node[1])
    if kind == "frac":
        top, bottom = _emit_sequence(node[1]), _emit_sequence(node[2])
        if not top or not bottom:
            raise MathError("A fraction needs both a numerator and a denominator.")
        return _group(f"({top})/({bottom})")
    if kind == "binom":
        return _group(f"(({_emit_sequence(node[1])}){ATOP}({_emit_sequence(node[2])}))")
    if kind == "sqrt":
        body = _emit_sequence(node[2])
        if not body:
            raise MathError("\\sqrt needs something under the root.")
        if node[1]:
            return _group(f"√({_emit_sequence(node[1])}&{body})")
        return _group(f"√({body})")
    if kind == "script":
        base = _emit(node[1])
        if not base:
            raise MathError("A subscript or superscript has nothing in front of it to attach to.")
        return _group(base + _script_part("_", node[2]) + _script_part("^", node[3]))
    if kind == "nary":
        return _emit_nary(node, "")
    if kind == "func":
        return _emit_function(node, "")
    if kind == "accent":
        body = _emit_sequence(node[2])
        if not body:
            raise MathError("An accent needs a letter to sit on.")
        if len(body) == 1:
            return body + node[1]
        return _group(f"({body}){node[1]}")
    if kind == "bar":
        return _group(f"{node[1]}({_emit_sequence(node[2])})")
    if kind == "delim":
        _, opener, closer, body = node
        inner = _emit_sequence(body)
        if opener and closer:
            return opener + inner + closer          # Word pairs them itself, also unlike ones such as [0,1)
        if opener:
            return _group(opener + inner + CLOSE_ANY + " ")
        if closer == ")":
            return _group(OPEN_ANY + " " + inner + ")")
        if closer in ("|", "‖"):
            return _group(OPEN_ANY + " " + inner + CLOSE_ANY + closer)
        if not closer:
            return _group(inner)
        raise MathError(_UNSUPPORTED.format("\\left. closed by " + closer))
    if kind == "matrix":
        _, env, rows = node
        body = "@".join("&".join(_emit_sequence(cell) for cell in row) for row in rows)
        if env == "cases":
            return _group("{" + ARRAY + "(" + body + ")" + CLOSE_ANY + " ")
        if env in MATRIX_ENVS:
            opener, closer = MATRIX_ENVS[env]
            return _group(opener + MATRIX + "(" + body + ")" + closer)
        return _group(ARRAY + "(" + body + ")")
    if kind == "over":
        return _group(_group(_emit_sequence(node[1])) + ABOVE + "(" + _emit_sequence(node[2]) + ")")
    if kind == "under":
        return _group(_group(_emit_sequence(node[1])) + BELOW + "(" + _emit_sequence(node[2]) + ")")
    if kind == "brace":
        return _group(f"{node[1]}({_emit_sequence(node[2])})")
    if kind == "styled":
        table = STYLES[node[1]]
        # Grouped: a bold letter at the very start of an equation would otherwise lose its style in Word
        return _group("".join(table.get(ch, ch) for ch in _emit_sequence(node[2])))
    raise MathError("Internal error: unknown node " + kind)


def strip_math_delimiters(src):
    """Drop $...$, $$...$$, \\(...\\) and \\[...\\] around a formula."""
    s = src.strip()
    for opener, closer in (("$$", "$$"), ("\\[", "\\]"), ("\\(", "\\)"), ("$", "$")):
        if len(s) >= len(opener) + len(closer) and s.startswith(opener) and s.endswith(closer):
            return s[len(opener):len(s) - len(closer)].strip()
    return s


def latex_to_unicodemath(src):
    """Translate one LaTeX formula. Raises MathError for anything outside the supported subset."""
    s = strip_math_delimiters(src)
    if not s:
        raise MathError("The formula is empty.")
    parser = _Parser(s)
    rows = parser.rows()
    if not parser.at_end():
        ch = parser.peek()
        if ch == "}":
            raise MathError("There is a } without a matching {.")
        raise MathError(f"Could not read the formula from: {parser.s[parser.i:parser.i + 20]!r}")
    if len(rows) > 1 or len(rows[0]) > 1:
        out = _emit(("matrix", "aligned", rows))
    else:
        out = _emit_sequence(rows[0][0])
    out = out.strip(" ")
    if not out:
        raise MathError("The formula is empty.")
    return out


def check_unicodemath(text):
    """Basic checks on a formula passed directly in Word's linear format."""
    s = text.strip()
    if not s:
        raise MathError("The formula is empty.")
    if "\r" in s or "\n" in s:
        raise MathError("A formula must be on one line. For several lines use an equation array: █(first line@second line).")
    for opener, closer in (("(", ")"), (GROUP_OPEN, GROUP_CLOSE)):
        if s.count(opener) != s.count(closer):
            raise MathError(f"Unbalanced {opener} {closer} in the formula: Word would leave part of it unbuilt.")
    return s


if __name__ == "__main__":
    import sys
    for line in sys.stdin:
        line = line.rstrip("\n")
        if line.strip():
            try:
                print(latex_to_unicodemath(line))
            except MathError as exc:
                print("ERROR:", exc)
