#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
实验五：语法制导翻译和中间代码生成
支持：
  Part 1 —— 含加减乘除、括号的赋值语句 → 四元组序列
  Part 2 —— 布尔表达式（and/or/not/relop）→ 四元组（回填法）
  Part 3 —— 含 if-else/while 条件语句    → 四元组（选做）
"""

import re, sys
from typing import List, Optional, Tuple

# ─────────────────────────────────────────────
# 词法分析
# ─────────────────────────────────────────────
TK_NUM   = 'NUM'
TK_ID    = 'ID'
TK_OP    = 'OP'
TK_REL   = 'RELOP'   # < <= > >= == !=
TK_AND   = 'AND'
TK_OR    = 'OR'
TK_NOT   = 'NOT'
TK_TRUE  = 'TRUE'
TK_FALSE = 'FALSE'
TK_IF    = 'IF'
TK_ELSE  = 'ELSE'
TK_WHILE = 'WHILE'
TK_LPAREN = '('
TK_RPAREN = ')'
TK_LBRACE = '{'
TK_RBRACE = '}'
TK_SEMI  = ';'
TK_ASSIGN = '='
TK_EOF   = 'EOF'

KEYWORDS = {'and': TK_AND, 'or': TK_OR, 'not': TK_NOT,
            'true': TK_TRUE, 'false': TK_FALSE,
            'if': TK_IF, 'else': TK_ELSE, 'while': TK_WHILE}

Token = Tuple[str, str]   # (type, value)

def tokenize(text: str) -> List[Token]:
    tokens: List[Token] = []
    i = 0
    text = text.strip()
    while i < len(text):
        if text[i].isspace():
            i += 1
            continue
        # 数字
        m = re.match(r'\d+(\.\d+)?', text[i:])
        if m:
            tokens.append((TK_NUM, m.group()))
            i += len(m.group())
            continue
        # 标识符 / 关键字
        m = re.match(r'[A-Za-z_]\w*', text[i:])
        if m:
            word = m.group()
            tk = KEYWORDS.get(word, TK_ID)
            tokens.append((tk, word))
            i += len(word)
            continue
        # 关系运算符（两字符优先）
        if text[i:i+2] in ('<=', '>=', '==', '!='):
            tokens.append((TK_REL, text[i:i+2]))
            i += 2
            continue
        if text[i] in '<>':
            tokens.append((TK_REL, text[i]))
            i += 1
            continue
        # 赋值 =
        if text[i] == '=':
            tokens.append((TK_ASSIGN, '='))
            i += 1
            continue
        # 算术运算符
        if text[i] in '+-*/':
            tokens.append((TK_OP, text[i]))
            i += 1
            continue
        # 括号 / 大括号 / 分号
        if text[i] == '(':
            tokens.append((TK_LPAREN, '('))
            i += 1; continue
        if text[i] == ')':
            tokens.append((TK_RPAREN, ')'))
            i += 1; continue
        if text[i] == '{':
            tokens.append((TK_LBRACE, '{'))
            i += 1; continue
        if text[i] == '}':
            tokens.append((TK_RBRACE, '}'))
            i += 1; continue
        if text[i] == ';':
            tokens.append((TK_SEMI, ';'))
            i += 1; continue
        raise SyntaxError(f"非法字符: {text[i]!r} (位置 {i})")
    tokens.append((TK_EOF, ''))
    return tokens


# ─────────────────────────────────────────────
# 符号表 & 四元组管理
# ─────────────────────────────────────────────
class SymbolTable:
    def __init__(self):
        self._table = {}      # name → index (1-based)
        self._names = []      # index → name
        self._temp_cnt = 0

    def lookup(self, name: str) -> int:
        if name not in self._table:
            self._names.append(name)
            self._table[name] = len(self._names)
        return self._table[name]

    def new_temp(self) -> Tuple[str, int]:
        self._temp_cnt += 1
        name = f't{self._temp_cnt}'
        idx = -(self._temp_cnt)   # 负数下标区分临时变量
        self._names_temp = getattr(self, '_names_temp', {})
        self._names_temp[idx] = name
        return name, idx

    def index(self, name: str) -> int:
        return self._table.get(name, 0)

    def dump(self):
        print("  符号表：")
        for i, n in enumerate(self._names, 1):
            print(f"    [{i}] {n}")


class QuadList:
    """四元组列表，支持回填"""
    def __init__(self):
        self._quads: List[Tuple] = []   # (op, arg1, arg2, result)

    def emit(self, op, arg1, arg2, result):
        self._quads.append((op, arg1, arg2, result))
        return len(self._quads) - 1    # 返回该条指令的编号

    def backpatch(self, lst: List[int], target: int):
        for idx in lst:
            op, a1, a2, _ = self._quads[idx]
            self._quads[idx] = (op, a1, a2, str(target))

    def next_quad(self) -> int:
        return len(self._quads)

    def dump(self, use_index=False, sym: Optional['SymbolTable'] = None):
        for i, (op, a1, a2, res) in enumerate(self._quads):
            print(f"  ({i}) ({op}, {a1}, {a2}, {res})")

    @property
    def quads(self):
        return self._quads


# ─────────────────────────────────────────────
# Part 1：赋值语句 / 算术表达式翻译
# ─────────────────────────────────────────────
# 文法（消除左递归后）：
#   S   → id = E
#   E   → T E'
#   E'  → + T E' | - T E' | ε
#   T   → F T'
#   T'  → * F T' | / F T' | ε
#   F   → ( E ) | id | num
#
# 语义规则：
#   F  → id        { F.place = id.place }
#   F  → num       { F.place = num.val }
#   F  → (E)       { F.place = E.place }
#   T  → F T'_inh  { T'.inh = F.place; T.place = T'.syn }
#   T' → * F T'1   { t = newtemp; emit(*,T'.inh,F.place,t);
#                     T'1.inh = t; T'.syn = T'1.syn }
#   T' → ε         { T'.syn = T'.inh }
#   E  → T E'_inh  { E'.inh = T.place; E.place = E'.syn }
#   E' → + T E'1   { t = newtemp; emit(+,E'.inh,T.place,t);
#                     E'1.inh = t; E'.syn = E'1.syn }
#   E' → ε         { E'.syn = E'.inh }
#   S  → id = E    { emit(=,E.place,0,id) }

class ArithParser:
    def __init__(self, tokens: List[Token], sym: SymbolTable, quad: QuadList):
        self.tokens = tokens
        self.pos = 0
        self.sym = sym
        self.quad = quad

    def peek(self) -> Token:
        return self.tokens[self.pos]

    def consume(self, expected_type=None, expected_val=None) -> Token:
        tok = self.tokens[self.pos]
        if expected_type and tok[0] != expected_type:
            raise SyntaxError(f"期望 {expected_type}，得到 {tok}")
        if expected_val and tok[1] != expected_val:
            raise SyntaxError(f"期望 '{expected_val}'，得到 {tok[1]!r}")
        self.pos += 1
        return tok

    def parse_S(self):
        """S → id = E"""
        id_tok = self.consume(TK_ID)
        self.consume(TK_ASSIGN)
        e_place = self.parse_E()
        self.sym.lookup(id_tok[1])   # 确保目标变量在符号表中
        self.quad.emit('=', e_place, '0', id_tok[1])

    def parse_E(self) -> str:
        """E → T E'"""
        t_place = self.parse_T()
        return self.parse_Eprime(t_place)

    def parse_Eprime(self, inh: str) -> str:
        """E' → + T E' | - T E' | ε"""
        tok = self.peek()
        if tok[0] == TK_OP and tok[1] in ('+', '-'):
            op = self.consume()[1]
            t_place = self.parse_T()
            tmp, _ = self.sym.new_temp()
            self.quad.emit(op, inh, t_place, tmp)
            return self.parse_Eprime(tmp)
        return inh   # ε

    def parse_T(self) -> str:
        """T → F T'"""
        f_place = self.parse_F()
        return self.parse_Tprime(f_place)

    def parse_Tprime(self, inh: str) -> str:
        """T' → * F T' | / F T' | ε"""
        tok = self.peek()
        if tok[0] == TK_OP and tok[1] in ('*', '/'):
            op = self.consume()[1]
            f_place = self.parse_F()
            tmp, _ = self.sym.new_temp()
            self.quad.emit(op, inh, f_place, tmp)
            return self.parse_Tprime(tmp)
        return inh

    def parse_F(self) -> str:
        """F → (E) | id | num"""
        tok = self.peek()
        if tok[0] == TK_LPAREN:
            self.consume(TK_LPAREN)
            e_place = self.parse_E()
            self.consume(TK_RPAREN)
            return e_place
        if tok[0] == TK_ID:
            self.consume()
            self.sym.lookup(tok[1])
            return tok[1]
        if tok[0] == TK_NUM:
            self.consume()
            return tok[1]
        raise SyntaxError(f"parse_F：意外的 token {tok}")


# ─────────────────────────────────────────────
# Part 2：布尔表达式翻译（回填法）
# ─────────────────────────────────────────────
# 文法：
#   B  → B or M B1 | B and M B1 | not B1 | ( B ) | E relop E | true | false
# 消除左递归（改写为右递归等价形式）：
#   B   → B1 B'
#   B'  → or M B B' | ε
#   B1  → B2 B1'
#   B1' → and M B1 B1' | ε
#   B2  → not B2 | ( B ) | E relop E | true | false
#   M   → ε { M.quad = nextquad }
#
# 语义属性：
#   truelist, falselist — 需回填的四元组编号列表
#   M.quad              — 标记当前四元组编号（用于回填目标）

def merge(l1: List[int], l2: List[int]) -> List[int]:
    return l1 + l2

class BoolParser:
    """独立的布尔表达式解析器（共用 sym / quad）"""
    def __init__(self, tokens: List[Token], sym: SymbolTable, quad: QuadList):
        self.tokens = tokens
        self.pos = 0
        self.sym = sym
        self.quad = quad

    def peek(self) -> Token:
        return self.tokens[self.pos]

    def consume(self, expected_type=None, expected_val=None) -> Token:
        tok = self.tokens[self.pos]
        if expected_type and tok[0] != expected_type:
            raise SyntaxError(f"期望 {expected_type}，得到 {tok}")
        if expected_val and tok[1] != expected_val:
            raise SyntaxError(f"期望 '{expected_val}'，得到 '{tok[1]}'")
        self.pos += 1
        return tok

    def parse_B(self):
        """B → B1 B'  return (truelist, falselist)"""
        tl, fl = self.parse_B1()
        return self.parse_Bprime(tl, fl)

    def parse_Bprime(self, inh_tl, inh_fl):
        """B' → or M B B' | ε"""
        if self.peek()[0] == TK_OR:
            self.consume(TK_OR)
            m_quad = self.quad.next_quad()    # M.quad
            self.quad.backpatch(inh_fl, m_quad)
            tl1, fl1 = self.parse_B1()
            tl1, fl1 = self.parse_Bprime(tl1, fl1)
            return merge(inh_tl, tl1), fl1
        return inh_tl, inh_fl

    def parse_B1(self):
        """B1 → B2 B1'"""
        tl, fl = self.parse_B2()
        return self.parse_B1prime(tl, fl)

    def parse_B1prime(self, inh_tl, inh_fl):
        """B1' → and M B1 B1' | ε"""
        if self.peek()[0] == TK_AND:
            self.consume(TK_AND)
            m_quad = self.quad.next_quad()
            self.quad.backpatch(inh_tl, m_quad)
            tl1, fl1 = self.parse_B2()
            tl1, fl1 = self.parse_B1prime(tl1, fl1)
            return tl1, merge(inh_fl, fl1)
        return inh_tl, inh_fl

    def parse_B2(self):
        """B2 → not B2 | (B) | E relop E | true | false"""
        tok = self.peek()
        if tok[0] == TK_NOT:
            self.consume(TK_NOT)
            tl, fl = self.parse_B2()
            return fl, tl    # 取反：交换 truelist / falselist
        if tok[0] == TK_LPAREN:
            self.consume(TK_LPAREN)
            tl, fl = self.parse_B()
            self.consume(TK_RPAREN)
            return tl, fl
        if tok[0] == TK_TRUE:
            self.consume()
            idx = self.quad.emit('j', '_', '_', '?')
            return [idx], []
        if tok[0] == TK_FALSE:
            self.consume()
            idx = self.quad.emit('j', '_', '_', '?')
            return [], [idx]
        # E relop E
        arith = ArithParser(self.tokens, self.sym, self.quad)
        arith.pos = self.pos
        e1 = arith.parse_E()
        self.pos = arith.pos
        rel_tok = self.consume(TK_REL)
        arith.pos = self.pos
        e2 = arith.parse_E()
        self.pos = arith.pos
        t_idx = self.quad.emit(f'j{rel_tok[1]}', e1, e2, '?')
        f_idx = self.quad.emit('j', '_', '_', '?')
        return [t_idx], [f_idx]


# ─────────────────────────────────────────────
# Part 3（选做）：条件语句 / while 语句翻译
# ─────────────────────────────────────────────
# 文法：
#   Stmt → if ( B ) { Stmt } ElsePart
#        | while ( B ) { Stmt }
#        | id = E ;
#   ElsePart → else { Stmt } | ε
#   Stmts → Stmt Stmts | ε
#
# 语义规则（简）：
#   if(B) S1 else S2:
#     backpatch(B.truelist, S1_start)
#     backpatch(B.falselist, S2_start)
#     goto after_S2
#   while(B) S:
#     loop_start = nextquad
#     backpatch(B.truelist, body_start)
#     backpatch(B.falselist, after_while)
#     at end of body: emit(j, _, _, loop_start)

class StmtParser:
    def __init__(self, tokens: List[Token], sym: SymbolTable, quad: QuadList):
        self.tokens = tokens
        self.pos = 0
        self.sym = sym
        self.quad = quad

    def peek(self) -> Token:
        return self.tokens[self.pos]

    def consume(self, expected_type=None, expected_val=None) -> Token:
        tok = self.tokens[self.pos]
        if expected_type and tok[0] != expected_type:
            raise SyntaxError(f"期望 {expected_type}，得到 {tok}")
        if expected_val and tok[1] != expected_val:
            raise SyntaxError(f"期望 '{expected_val}'，得到 '{tok[1]}'")
        self.pos += 1
        return tok

    def parse_stmts(self) -> List[int]:
        """Stmts → Stmt Stmts | ε，返回 nextlist（跳出当前块的 j 链）"""
        nextlist = []
        while self.peek()[0] not in (TK_EOF, TK_RBRACE):
            nl = self.parse_stmt()
            nextlist = merge(nextlist, nl)
        return nextlist

    def parse_stmt(self) -> List[int]:
        tok = self.peek()
        if tok[0] == TK_IF:
            return self.parse_if()
        if tok[0] == TK_WHILE:
            return self.parse_while()
        if tok[0] == TK_ID:
            return self.parse_assign()
        raise SyntaxError(f"parse_stmt：意外 token {tok}")

    def parse_assign(self) -> List[int]:
        arith = ArithParser(self.tokens, self.sym, self.quad)
        arith.pos = self.pos
        arith.parse_S()
        self.pos = arith.pos
        self.consume(TK_SEMI)
        return []

    def parse_if(self) -> List[int]:
        self.consume(TK_IF)
        self.consume(TK_LPAREN)
        bp = BoolParser(self.tokens, self.sym, self.quad)
        bp.pos = self.pos
        tl, fl = bp.parse_B()
        self.pos = bp.pos
        self.consume(TK_RPAREN)
        # 真分支
        true_start = self.quad.next_quad()
        self.quad.backpatch(tl, true_start)
        self.consume(TK_LBRACE)
        s1_next = self.parse_stmts()
        self.consume(TK_RBRACE)
        # 跳过 else 的 goto
        goto_idx = self.quad.emit('j', '_', '_', '?')
        # else 分支
        false_start = self.quad.next_quad()
        self.quad.backpatch(fl, false_start)
        if self.peek()[0] == TK_ELSE:
            self.consume(TK_ELSE)
            self.consume(TK_LBRACE)
            s2_next = self.parse_stmts()
            self.consume(TK_RBRACE)
            return merge(merge(s1_next, [goto_idx]), s2_next)
        else:
            return merge(merge(s1_next, [goto_idx]), fl)

    def parse_while(self) -> List[int]:
        loop_start = self.quad.next_quad()
        self.consume(TK_WHILE)
        self.consume(TK_LPAREN)
        bp = BoolParser(self.tokens, self.sym, self.quad)
        bp.pos = self.pos
        tl, fl = bp.parse_B()
        self.pos = bp.pos
        self.consume(TK_RPAREN)
        body_start = self.quad.next_quad()
        self.quad.backpatch(tl, body_start)
        self.consume(TK_LBRACE)
        s_next = self.parse_stmts()
        self.consume(TK_RBRACE)
        # 回到循环头
        self.quad.emit('j', '_', '_', str(loop_start))
        self.quad.backpatch(s_next, loop_start)
        return fl   # falselist：循环后继


# ─────────────────────────────────────────────
# 格式化输出
# ─────────────────────────────────────────────
def print_quads(quads, title="四元组序列"):
    print(f"\n{title}")
    print("─" * 40)
    for i, (op, a1, a2, res) in enumerate(quads):
        print(f"  ({i:2d})  ({op}, {a1}, {a2}, {res})")
    print("─" * 40)


# ─────────────────────────────────────────────
# 主测试入口
# ─────────────────────────────────────────────
def run_arith(expr: str):
    sym = SymbolTable()
    quad = QuadList()
    tokens = tokenize(expr)
    p = ArithParser(tokens, sym, quad)
    p.parse_S()
    return sym, quad

def run_bool(expr: str):
    sym = SymbolTable()
    quad = QuadList()
    tokens = tokenize(expr)
    bp = BoolParser(tokens, sym, quad)
    bp.pos = 0
    tl, fl = bp.parse_B()
    return sym, quad, tl, fl

def run_stmt(code: str):
    sym = SymbolTable()
    quad = QuadList()
    tokens = tokenize(code)
    sp = StmtParser(tokens, sym, quad)
    nextlist = sp.parse_stmts()
    return sym, quad, nextlist


if __name__ == '__main__':
    SEP = "=" * 50

    # ── Part 1: 算术表达式赋值语句 ────────────
    print(SEP)
    print("Part 1：赋值语句 → 四元组")
    print(SEP)

    cases_arith = [
        "a=b+c*e/g",
        "x=a+b*c-d/e",
        "result=(a+b)*(c-d)",
        "y=a*b+c*d+e*f",
        "z=((a+b)*c-d)/e+f",
    ]
    for expr in cases_arith:
        print(f"\n输入: {expr}")
        sym, quad = run_arith(expr)
        print_quads(quad.quads)

    # ── Part 2: 布尔表达式（回填） ────────────
    print("\n")
    print(SEP)
    print("Part 2：布尔表达式 → 四元组（回填法）")
    print(SEP)

    cases_bool = [
        ("a < b",          "简单关系表达式"),
        ("a < b and c > d","and 表达式"),
        ("a < b or c > d", "or 表达式"),
        ("not a < b",      "not 表达式"),
        ("a<b and c>d or e<f", "复合布尔表达式"),
        ("not (a<b or c>d) and e<=f", "含 not 和括号"),
    ]
    for expr, desc in cases_bool:
        print(f"\n输入: {expr}  （{desc}）")
        sym, quad, tl, fl = run_bool(expr)
        print_quads(quad.quads)
        print(f"  truelist  = {tl}")
        print(f"  falselist = {fl}")

    # ── Part 3: 条件/循环语句 ─────────────────
    print("\n")
    print(SEP)
    print("Part 3（选做）：条件语句 / while 语句 → 四元组")
    print(SEP)

    prog_if = "if (a < b) { x = a + 1 ; } else { x = b + 2 ; }"
    print(f"\n输入:\n  {prog_if}")
    sym, quad, nl = run_stmt(prog_if)
    print_quads(quad.quads, "if-else 语句四元组")
    print(f"  nextlist = {nl}")

    prog_while = "while (i < n) { i = i + 1 ; }"
    print(f"\n输入:\n  {prog_while}")
    sym, quad, nl = run_stmt(prog_while)
    print_quads(quad.quads, "while 语句四元组")
    print(f"  nextlist = {nl}")

    prog_nested = (
        "if (a < b and c > 0) { "
        "  result = a * b ; "
        "} else { "
        "  result = b + c ; "
        "}"
    )
    print(f"\n输入:\n  {prog_nested}")
    sym, quad, nl = run_stmt(prog_nested)
    print_quads(quad.quads, "复合条件 if-else 四元组")
    print(f"  nextlist = {nl}")

    print(f"\n{SEP}")
    print("所有测试执行完毕。")
    print(SEP)
