#!/usr/bin/env python3
"""Test workbooks shaped like the CAPA workbook after CAPA_Setup (macro v13).

  A_formulas.xlsx  formulas in the cells (CAPA_RULES mode FORMULAS), values calculated by LibreOffice
  A_values.xlsx    formulas as rules in CAPA_RULES (mode VALUES), the values of A_formulas in the cells
  B_formulas.xlsx  the same with other inputs (plan, coefficients, PART METHOD, ...)
  AB_values.xlsx   the inputs of B with the (old) values of A: a recalculation must give B
  expected_B.json  every value of B (formula cells from LibreOffice, 1-2 CH:AGS and ST_SUM from the oracle)

1-2 CH:AGS and ST_SUM are what CAPA_Recalc writes; they are computed here by an independent
Python version of BuildTotals / BuildQtyTotals (the oracle). Everything else is calculated by
LibreOffice from the formulas.
"""
import json
import os
import random
import re
import shutil
import subprocess
import sys
import zipfile

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fixtures')

FIRST_ROW, COEF_ROW, HDR_ROW = 7, 3, 4
COL_PN, COL_OD, COL_ST, N_PROC = 6, 12, 21, 65
COL_PLAN, BLOCK, N_MONTH = 86, 66, 12
COL_QTY, COL_CHILD, COL_MK_FM, COL_MK_BD, COL_P, COL_PCS = 879, 8, 18, 19, 16, 11
PLAN_ROW, PLAN_COL_PN, PLAN_COL_M1 = 4, 3, 7
CAPA_COL, CAPA_FLAG = 73, 8
FM_P1, FM_P2, FM_STOP, BD_P = 8, 15, 11, 16
COL_CH_FM, COL_CH_BD = 13, 14
BULLET = '●'
MONTHS = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
S_PLAN, S_M, S12, S_SUM, S13 = '1-1.Plan', 'PART METHOD', '1-2.ITEM', 'ST_SUM', '1-3.CAPA'
S_BP, S_IC, S_TEST, S_RULES = 'BY PROCESS', 'BY ITEM CODE', 'TEST', 'CAPA_RULES'
SHEETS = [S_PLAN, S_M, S12, S_SUM, S13, S_BP, S_IC, S_TEST, S_RULES]
PROC_NAMES = {8: 'Reduce', 9: 'Expand', 10: 'Dimple', 11: 'Stopper', 12: 'Drill', 13: 'Spin', 14: 'Pinch Off',
              15: 'Press Joint', 16: 'Bending Pipe'}


class Err(str):
    """An error value (#N/A, ...)."""


# ---- values as in the VBA module ----------------------------------------------------
PLAIN = re.compile(r'^[ \t\n\r\f\v]*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?[ \t\n\r\f\v]*$')


def is_plain_number(s):
    if not s:
        return False
    if s[0].isalpha() and s[0].isascii():
        return False
    return PLAIN.match(s) is not None


def to_number(x):
    """(value, error) like ToNumber of modCapaCalc."""
    if x is None:
        return 0.0, None
    if isinstance(x, Err):
        return 0.0, x
    if isinstance(x, bool):
        return (1.0 if x else 0.0), None
    if isinstance(x, (int, float)):
        return float(x), None
    s = str(x)
    if s.replace(' ', ' ').strip(' ') == '':
        return 0.0, None
    if is_plain_number(s):
        return float(s), None
    return 0.0, Err('#VALUE!')


def k15(x):
    return '0' if x == 0 else format(x, '.15g')


def crit_key(x):
    if x is None:
        return ''
    if isinstance(x, Err):
        return 'E' + x
    if isinstance(x, bool):
        return 'B' + str(x)
    if isinstance(x, str):
        if x == '':
            return ''
        if is_plain_number(x):
            return 'N' + k15(float(x))
        return 'S' + x.upper()
    return 'N' + k15(float(x))


def vb_round(x):
    return round(x)   # Python rounds half to even like VBA


def cnc2_shots(n, pcs):
    if abs(n - vb_round(n)) > 1e-9 or n > 1000000:
        return n / 2
    p, e = to_number(pcs)
    if e is not None or p < 1 or p > 1000000 or abs(p - vb_round(p)) > 1e-9:
        p = 1
    nn, pc = int(vb_round(n)), int(vb_round(p))
    q = nn // pc
    rm = nn - q * pc
    return float((pc - rm) * ((q + 1) // 2) + rm * ((q + 2) // 2))


# ---- inputs ------------------------------------------------------------------------------
def make_inputs(variant, clean=False):
    rnd = random.Random(7)
    specs = [6.35, 9.52, 12.7, 'SPEC-A', 15.88]
    inp = {'variant': variant}
    # 1-1.Plan: duplicates, a lower case duplicate, a number P/N, an error, a text in a quantity cell
    pns = ['ASSY-%03d' % i for i in range(1, 31)]
    plan = []
    for pn in pns:
        plan.append([pn] + [float(rnd.randint(0, 400)) for _ in range(12)])
    plan.append(['ASSY-001'] + [float(rnd.randint(0, 50)) for _ in range(12)])
    plan.append(['assy-002'] + [10.0] * 12)
    plan.append([12345] + [float(rnd.randint(1, 90)) for _ in range(12)])
    plan.append([None] + [99.0] * 12)
    if not clean:
        plan[4][12] = Err('#N/A')               # ASSY-005 DEC
        plan[6][3] = 'n/a'                      # ignored by SUMIF
    if variant == 'B':
        r2 = random.Random(99)
        for row in plan:
            for m in range(1, 13):
                if isinstance(row[m], float) and r2.random() < 0.5:
                    row[m] = float(r2.randint(0, 500))
        plan[4][12] = 17.0                      # the error is gone
        if not clean:
            plan[10][5] = Err('#DIV/0!')        # ASSY-011 APR
    inp['plan'] = plan

    items = []
    n_items = 64
    for i in range(n_items):
        it = {}
        it['pn'] = rnd.choice(pns[:28]) if i % 11 else ('ASSY-999' if i % 22 == 0 else '12345')
        if i == 5:
            it['pn'] = 'ASSY-005'
        if i == 10:
            it['pn'] = 'ASSY-011'
        it['child'] = 'CH-%03d' % (i % 25)
        it['pcs'] = float(rnd.choice([1, 1, 2, 3, 4]))
        od = specs[i % len(specs)]
        if od == 15.88 and i > 20 and i % 2 == 0:
            od = '15.88'                        # number text: the same OD group
        if i == 5:
            od = 'SPEC-A'
        it['od'] = od
        it['p'] = 0.0 if i % 9 == 4 else float(rnd.choice([1, 2]))
        it['mk_fm'] = 'FM' if i % 3 else 0.0
        it['mk_bd'] = 'BD' if i % 4 else 0.0
        st = {}
        for k in range(1, N_PROC + 1):
            if rnd.random() < 0.35 or (FM_P1 <= k <= BD_P and rnd.random() < 0.6):
                st[k] = round(rnd.uniform(1, 60), 2) if rnd.random() < 0.7 else float(rnd.randint(1, 40))
        it['st'] = st
        items.append(it)
    items[3]['st'][3] = '12'                    # number text
    items[4]['st'][5] = '   '                   # spaces only = blank
    if not clean:
        items[7]['st'][40] = 'abc'              # #VALUE! for the whole row
    items[7]['od'] = 12.7
    inp['items'] = items

    coef = []
    for m in range(12):
        coef.append([1.0 if rnd.random() < 0.7 else rnd.choice([1.1, 0.9, 1.25]) for _ in range(N_PROC)])
    if variant == 'B':
        coef[2][7] = 1.5
        coef[5][15] = 0.8
    inp['coef'] = coef

    methods = [
        ('CH-001', 'FM', 'FMCNC1', None),
        ('CH-002', 'FM', 'FMCNC2', 5.0),
        ('CH-003', 'BD', 'HAND', 12.0),
        ('CH-004', 'BD', 'HAND', None),
        ('CH-005', 'FM', 'CNC2', None),         # name of v10-v12
        ('CH-006', 'BD', 'MANUAL', 9.0),        # name of v10-v12
        ('CH-007', 'FM ', 'fmcnc 1', '  '),     # spaces / case; time of only spaces = none
        ('CH-099', 'FM', 'FMCNC1', None),       # not in sheet 1-2
        ('CH-008', 'FM', 'XYZ', None),          # unknown method
        ('CH-002', 'BD', 'HAND', None),         # the same part twice
    ]
    if variant == 'B':
        methods[2] = ('CH-003', 'BD', 'HAND', 15.0)
        methods.append(('CH-010', 'FM', 'FMCNC2', None))
    inp['methods'] = methods
    inp['ms'] = {'D4': 'FMCNC1', 'E4': 7.0, 'D5': 'HAND', 'D6': 'FMCNC2', 'E6': 8.0 if variant == 'B' else 7.5}

    machines = []
    for s in specs:
        for kind in ['NORMAL', 'NORMAL', 'FMCNC1', 'FMCNC2', 'HAND', 'OTHER']:
            flags = set()
            for k in range(1, N_PROC + 1):
                if kind in ('FMCNC1', 'FMCNC2'):
                    if FM_P1 <= k <= FM_P2 and rnd.random() < 0.8:
                        flags.add(k)
                elif kind == 'HAND':
                    if k == BD_P:
                        flags.add(k)
                elif rnd.random() < 0.6:
                    flags.add(k)
            machines.append({'kind': kind, 'spec': s, 'flags': flags, 'e': float(rnd.choice([8, 16]))})
    machines[3]['kind'] = ' FMCNC1'             # almost the Kind of D4: another machine for the formulas
    inp['machines'] = machines
    return inp


# ---- oracle: BuildTotals / BuildQtyTotals / WriteMethodMarks --------------------------------
def oracle(inp):
    nrow = len(inp['items'])
    ncol = N_MONTH * BLOCK
    plan_keys, plan_v, plan_e = {}, [], []
    for row in inp['plan']:
        key = crit_key(row[0])
        if not key:
            continue
        if key not in plan_keys:
            plan_keys[key] = len(plan_v)
            plan_v.append([0.0] * 12)
            plan_e.append([None] * 12)
        idx = plan_keys[key]
        for m in range(12):
            if plan_e[idx][m] is None:
                q = row[m + 1]
                if isinstance(q, Err):
                    plan_e[idx][m] = q
                elif isinstance(q, float):
                    plan_v[idx][m] += q
    # PART METHOD
    meth = {}
    bad = 0
    for child, shop, how, t in inp['methods']:
        key = crit_key(child)
        shop = shop.strip(' ').upper()
        how = how.replace(' ', '').upper()
        it = list(meth.get(key, [0, None, False, None]))
        mode = 0
        if shop == 'FM':
            if how in ('FMCNC1', 'CNC1', 'CNC'):
                mode = 1
            elif how in ('FMCNC2', 'CNC2'):
                mode = 2
        tt = None if t is None or (isinstance(t, str) and t.strip(' ') == '') else to_number(t)[0]
        if mode:
            it[0], it[1] = mode, tt
            meth[key] = it
        elif shop == 'BD' and how in ('HAND', 'MANUAL'):
            it[2], it[3] = True, tt
            meth[key] = it
        else:
            bad += 1
    items = inp['items']
    fm_mode, bd_m, bd_t, ovr, fm_sec = [0] * nrow, [False] * nrow, [None] * nrow, [False] * nrow, [0.0] * nrow
    sec = {1: inp['ms']['E4'], 2: inp['ms']['E6']}
    for r, itm in enumerate(items):
        key = crit_key(itm['child'])
        if key in meth:
            it = meth[key]
            fm_mode[r], bd_m[r], bd_t[r] = it[0], it[2], it[3]
            if fm_mode[r]:
                fm_sec[r] = sec[fm_mode[r]] if it[1] is None else it[1]
            ovr[r] = fm_mode[r] > 0 or (bd_m[r] and bd_t[r] is not None)
    marks = []
    for r, itm in enumerate(items):
        a, b = itm['mk_fm'], itm['mk_bd']
        if isinstance(a, str) and a.strip(' ').upper() in ('FM', 'CNC1', 'CNC2', 'FMCNC1', 'FMCNC2'):
            a = ['FM', 'FMCNC1', 'FMCNC2'][fm_mode[r]]
        if isinstance(b, str) and b.strip(' ').upper() in ('BD', 'HAND'):
            b = 'HAND' if bd_m[r] else 'BD'
        marks.append((a, b))
    coef_v, coef_e, coef_err = [0.0] * (ncol + 1), [None] * (ncol + 1), [False] * 13
    for j in range(1, ncol + 1):
        m, k = (j - 1) // BLOCK, (j - 1) % BLOCK
        raw = None if k == 0 else inp['coef'][m][k - 1]
        v, e = to_number(raw)
        coef_v[j], coef_e[j] = v, e
        if e is not None:
            coef_err[m + 1] = True
    plan_row = []
    for itm in items:
        key = crit_key(itm['pn'])
        if key and key in plan_keys:
            idx = plan_keys[key]
            plan_row.append([plan_e[idx][m] if plan_e[idx][m] is not None else plan_v[idx][m] for m in range(12)])
        else:
            plan_row.append([0.0] * 12)
    grp_keys, grp_of, grp_val, grp_fm, grp_bd = {}, [], [], [], []
    for r, itm in enumerate(items):
        key = crit_key(itm['od']) + '\t' + str(fm_mode[r]) + ('M' if bd_m[r] else 'S')
        if key not in grp_keys:
            grp_keys[key] = len(grp_val)
            grp_val.append(itm['od'])
            grp_fm.append(['STD', 'FMCNC1', 'FMCNC2'][fm_mode[r]])
            grp_bd.append('HAND' if bd_m[r] else 'STD')
        grp_of.append(grp_keys[key])
    ngrp = len(grp_val)
    st_raw = [[itm['st'].get(k) for k in range(1, N_PROC + 1)] for itm in items]
    st_v, row_err = [[0.0] * N_PROC for _ in items], [False] * nrow
    for r in range(nrow):
        for k in range(N_PROC):
            v, e = to_number(st_raw[r][k])
            st_v[r][k] = v
            if e is not None:
                row_err[r] = True
    pair_f = [1.0] * nrow
    for r in range(nrow):
        if fm_mode[r] == 2:
            sv = st_v[r][FM_P1 - 1] + st_v[r][FM_P1] + st_v[r][FM_STOP - 1]
            if sv > 0:
                pair_f[r] = cnc2_shots(sv, items[r]['pcs']) / sv
    tot = [[0.0] * (ncol + 1) for _ in range(ngrp)]
    tot_e = [[None] * (ncol + 1) for _ in range(ngrp)]
    region = {}
    for m in range(1, 13):
        c = (m - 1) * BLOCK + 1
        for r in range(nrow):
            region[(FIRST_ROW + r, COL_PLAN + c - 1)] = plan_row[r][m - 1]
        for k in range(1, N_PROC + 1):
            j = c + k
            cv = coef_v[j]
            for r in range(nrow):
                pl = plan_row[r][m - 1]
                pe = pl if isinstance(pl, Err) else None
                pv = 0.0 if pe is not None else pl
                slow = row_err[r] or coef_err[m] or pe is not None or ovr[r]
                g = grp_of[r]
                if slow:
                    use_ov, ov, e = False, 0.0, None
                    sv = 0.0
                    if ovr[r]:
                        if fm_mode[r] > 0 and FM_P1 <= k <= FM_P2:
                            use_ov, ov = True, fm_sec[r]
                            if k <= FM_P1 + 1 or k == FM_STOP:
                                ov = ov * pair_f[r]
                            sv, e = to_number(st_raw[r][k - 1])
                            if e is None:
                                e = pe
                        elif k == BD_P and bd_m[r] and bd_t[r] is not None:
                            use_ov, ov = True, bd_t[r]
                            sv, e = to_number(items[r]['pcs'])
                            if e is None:
                                e = pe
                    if not use_ov:
                        sv, e = to_number(st_raw[r][k - 1])
                        if e is None:
                            e = coef_e[j]
                        if e is None:
                            e = pe
                    if e is None:
                        x = sv * ov * pv if use_ov else sv * cv * pv
                        tot[g][j] += x
                        region[(FIRST_ROW + r, COL_PLAN + j - 1)] = x
                    else:
                        if tot_e[g][j] is None:
                            tot_e[g][j] = e
                        region[(FIRST_ROW + r, COL_PLAN + j - 1)] = e
                else:
                    if pv != 0:
                        x = st_v[r][k - 1] * cv * pv
                        tot[g][j] += x
                    else:
                        x = 0.0
                    region[(FIRST_ROW + r, COL_PLAN + j - 1)] = x
    stsum = {}
    for g in range(ngrp):
        row = FIRST_ROW + g
        stsum[(row, COL_OD)] = grp_val[g]
        stsum[(row, COL_CH_FM)] = grp_fm[g]
        stsum[(row, COL_CH_BD)] = grp_bd[g]
        for j in range(1, ncol + 1):
            if (j - 1) % BLOCK != 0:
                stsum[(row, COL_PLAN + j - 1)] = tot_e[g][j] if tot_e[g][j] is not None else tot[g][j]
    # Q'TY AGU:AHF =IF($P>0, Plan*$K, 0) and its totals per ChildP/N
    qty = {}
    for r, itm in enumerate(items):
        for m in range(12):
            pl = plan_row[r][m]
            if itm['p'] > 0:
                q = pl if isinstance(pl, Err) else pl * itm['pcs']
            else:
                q = 0.0
            qty[(FIRST_ROW + r, COL_QTY + m)] = q
    q_keys, q_tot, q_val = {}, [], []
    for r, itm in enumerate(items):
        vals = [qty[(FIRST_ROW + r, COL_QTY + m)] for m in range(12)]
        if not any((isinstance(v, Err) or (isinstance(v, float) and v != 0)) for v in vals):
            continue
        key = crit_key(itm['child'])
        if key not in q_keys:
            q_keys[key] = len(q_val)
            q_val.append(itm['child'])
            q_tot.append([0.0] * 12)
        t = q_tot[q_keys[key]]
        for m in range(12):
            if isinstance(t[m], Err):
                continue
            if isinstance(vals[m], Err):
                t[m] = vals[m]
            else:
                t[m] += vals[m]
    for g, cv in enumerate(q_val):
        stsum[(FIRST_ROW + g, COL_CHILD)] = cv
        for m in range(12):
            stsum[(FIRST_ROW + g, COL_QTY + m)] = q_tot[g][m]
    return {'region': region, 'stsum': stsum, 'marks': marks, 'qty': qty, 'bad': bad}


# ---- formulas (R1C1) ----------------------------------------------------------------------------
def rc_col(v, rel, base):
    return v + base if rel else v


REF_TOKEN = re.compile(
    r'"(?:[^"]|"")*"'
    r"|'(?:[^']|'')*'!"
    r'|[A-Za-z_][A-Za-z0-9_.]*(?=\()'
    r'|[A-Za-z_][A-Za-z0-9_]*!'
    r'|R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?(?::R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?)?'
    r'|C(?:\[-?\d+\]|\d+)(?::C(?:\[-?\d+\]|\d+))?'
    r'|R(?:\[-?\d+\]|\d+)(?::R(?:\[-?\d+\]|\d+))?'
    r'|.', re.S)
PART = re.compile(r'^(R(\[-?\d+\]|\d+)?)?(C(\[-?\d+\]|\d+)?)?$')


def off(s, base):
    if s is None or s == '':
        return base, False
    if s.startswith('['):
        return base + int(s[1:-1]), False
    return int(s), True


def r1c1_to_a1(f, r0, c0):
    out = []
    for m in REF_TOKEN.finditer(f):
        t = m.group(0)
        if (t.startswith('R') or t.startswith('C')) and not t.endswith('!') and len(t) > 0 and PART.match(t.split(':')[0]):
            parts = t.split(':')
            conv = []
            for p in parts:
                pm = PART.match(p)
                has_r, has_c = pm.group(1) is not None, pm.group(3) is not None
                rr, ra = off(pm.group(2), r0)
                cc, ca = off(pm.group(4), c0)
                if has_r and has_c:
                    conv.append(('$' if ca else '') + get_column_letter(cc) + ('$' if ra else '') + str(rr))
                elif has_c:
                    conv.append(('$' if ca else '') + get_column_letter(cc))
                else:
                    conv.append(('$' if ra else '') + str(rr))
            if len(conv) == 1 and not (PART.match(parts[0]).group(1) and PART.match(parts[0]).group(3)):
                conv = conv * 2
            out.append(':'.join(conv))
        else:
            out.append(t)
    return ''.join(out)


def formulas(inp):
    """[(sheet, row, col, r1c1)]"""
    fs = []
    nrow = len(inp['items'])
    for r in range(FIRST_ROW, FIRST_ROW + nrow):
        for m in range(12):
            col = COL_QTY + m
            pc = COL_PLAN + m * BLOCK
            fs.append((S12, r, col, '=IF(RC16>0,RC[%d]*RC11,0)' % (pc - col)))
    k1, k2, kb = "'PART METHOD'!R4C4", "'PART METHOD'!R6C4", "'PART METHOD'!R5C4"
    nm = len(inp['machines'])
    for r in range(FIRST_ROW, FIRST_ROW + nm):
        for m in range(12):
            for k in range(1, N_PROC + 1):
                col = CAPA_COL + m * BLOCK + k - 1
                fl = CAPA_FLAG + k - 1
                cnt = 'COUNTIFS(C7,RC7,C%d,"%s",C1,R5C,C2,' % (fl, BULLET)
                if FM_P1 <= k <= FM_P2:
                    f = ('=IF(RC%d="%s",SUMIFS(ST_SUM!C[14],ST_SUM!C12,RC7,ST_SUM!C13,IF(RC2=%s,"FMCNC1",IF(RC2=%s,"FMCNC2",'
                         '"STD")))/IF(RC2=%s,%s%s),IF(RC2=%s,%s%s),%s"<>"&%s,C2,"<>"&%s)))/R3C,0)/3600/RC5') % (
                        fl, BULLET, k1, k2, k1, cnt, k1, k2, cnt, k2, cnt, k1, k2)
                elif k == BD_P:
                    f = ('=IF(RC%d="%s",SUMIFS(ST_SUM!C[14],ST_SUM!C12,RC7,ST_SUM!C14,IF(RC2=%s,"HAND","STD"))/%s'
                         'IF(RC2=%s,%s,"<>"&%s))/R3C,0)/3600/RC5') % (fl, BULLET, kb, cnt, kb, kb, kb)
                else:
                    f = ('=IF(RC%d="%s",SUMIFS(ST_SUM!C[14],ST_SUM!C12,RC7)/COUNTIFS(C7,RC7,C%d,"%s",C1,R5C)/R3C,0)'
                         '/3600/RC5') % (fl, BULLET, fl, BULLET)
                fs.append((S13, r, col, f))
            col = CAPA_COL + m * BLOCK + N_PROC
            fs.append((S13, r, col, '=SUM(RC[-65]:RC[-1])'))
    # BY PROCESS: load per Kind; rows 5.. (C = Kind; the FMCNC rows read PART METHOD D4 / D6)
    bp_rows = bp_layout()
    for i, (k, kind) in enumerate(bp_rows):
        r = 5 + i
        if kind == '=D4':
            fs.append((S_BP, r, 3, "='PART METHOD'!R4C4"))
        elif kind == '=D6':
            fs.append((S_BP, r, 3, "='PART METHOD'!R6C4"))
        for m in range(12):
            c13 = CAPA_COL + m * BLOCK + k - 1
            fs.append((S_BP, r, 4 + m, "=SUMIFS('1-3.CAPA'!C%d,'1-3.CAPA'!C2,RC3)" % c13))
            fs.append((S_BP, r, 17 + m, "=IFERROR(SUMIFS('1-3.CAPA'!C%d,'1-3.CAPA'!C2,RC3)/COUNTIFS('1-3.CAPA'!C2,RC3,"
                                         "'1-3.CAPA'!C%d,\"%s\"),0)" % (c13, CAPA_FLAG + k - 1, BULLET)))
        fs.append((S_BP, r, 29, '=SUM(RC4:RC15)'))
        fs.append((S_BP, r, 30, '=IF(RC29>100,"OVER",IF(RC29>0,INT(RC29*100%)&" h","-"))'))
        fs.append((S_BP, r, 31, '=COUNT(RC4:RC15)'))
    # BY ITEM CODE: Q'TY per ChildP/N (ST_SUM H / AGU:AHF), rows 4..
    children = item_code_list(inp)
    for i in range(len(children)):
        r = 4 + i
        for m in range(12):
            fs.append((S_IC, r, 8 + m, '=SUMIF(ST_SUM!C8,RC7,ST_SUM!C%d)' % (COL_QTY + m)))
        fs.append((S_IC, r, 20, '=SUM(RC8:RC19)'))
        fs.append((S_IC, r, 21, '=IFERROR(VLOOKUP(RC7,%s!R7C8:R200C12,2,FALSE),"-")' % "'1-2.ITEM'"))
    # TEST: engine semantics
    for i, f in enumerate(TEST_FORMULAS):
        fs.append((S_TEST, 1 + i, 3, f))
    return fs


def bp_layout():
    rows = []
    for k in [1, 3, 8, 9, 11, 15, 16, 20, 40]:
        for kind in ['NORMAL', '=D4', '=D6', 'HAND', 'OTHER']:
            rows.append((k, kind))
    return rows


def item_code_list(inp):
    seen = []
    for it in inp['items']:
        if it['child'] not in seen:
            seen.append(it['child'])
    return seen + ['CH-404']


TEST_FORMULAS = [
    '=R1C1*2+1', '=-2^2', '=2^-1', '=R1C1&"x"', '="v"&0.1+0.2', '=IF(R1C1>3,"big","small")',
    '=IF(R2C1="ABC",1,0)', '=R3C1=""', '=R3C1=0', '=R3C1+1', '=IFERROR(1/R7C1,"div0")', '=IFERROR(R9C1,-1)',
    '=NA()', '=SUM(R1C1:R8C1)', '=COUNT(R1C1:R10C1)', '=SUMIF(R1C1:R10C1,">1")', '=COUNTIF(R1C1:R10C1,"abc")',
    '=COUNTIF(R1C1:R10C1,"a*")', '=COUNTIF(R1C1:R10C1,"<>5")', '=VLOOKUP("B",R1C5:R5C6,2,FALSE)',
    '=VLOOKUP(3,R1C5:R5C6,2,0)', '=VLOOKUP("zz",R1C5:R5C6,2,FALSE)', '=INT(R6C1)', '=MID("hello world",7,5)',
    '=MID(R1C1*1000,2,2)', '=IF(R4C1,"yes","no")', '=R4C1+1', '=10%', '=R1C1*50%',
    '=COUNTIFS(R1C1:R10C1,">=0",R1C2:R10C2,"x")', '=SUMIFS(R1C1:R10C1,R1C2:R10C2,"x",R1C1:R10C1,"<5")',
    '=IF(1<"a",1,0)', '="b">"a"', '=IF(R9C1=1,1,2)', '=IFERROR(IF(R9C1=1,1,2),"e")', '=R1C1/0',
    '=INT(7.9)+INT(-0.1)', '=(1-0.9)-0.1', '=COUNTIF(R1C1:R10C1,"")', '=COUNTIF(R1C1:R10C1,"<>")',
    '=SUMIF(R1C2:R10C2,"x",R1C1:R10C1)', '=SUMIF(R1C2:R10C2,"x",R1C1)', '=R[-1]C+R[-2]C', '=IF(R1C1=5,)',
    '=VLOOKUP("c",R1C5:R5C6,2,)', '=MID("abc",5,2)', '=R1C1-R1C1*0.9999999999999999',
    '=IFERROR(VLOOKUP(R1C1,R1C5:R5C6,2,FALSE),"none")', '=SUM(R1C1:R2C1,R6C1,4)',
]
TEST_INPUTS = [5.0, 'abc', None, True, '7', -2.5, 0.0, 'ABC', Err('#N/A'), 3.0]
TEST_B = ['x', 'x', 'y', None, 'x', None, 'x', None, None, None]
TEST_TABLE = [('a', 1.0), ('b', 2.0), (3.0, 'three'), ('C', 4.0), (None, 5.0)]


# ---- writing the workbook -------------------------------------------------------------------------
def put(ws, r, c, v):
    cell = ws.cell(row=r, column=c)
    if isinstance(v, Err):
        cell.value = str(v)
        cell.data_type = 'e'
    else:
        cell.value = v


def rects_of(cells):
    """cells [(r, c)] -> 'r1 c1 r2 c2,...' (runs down the columns)"""
    by_col = {}
    for r, c in cells:
        by_col.setdefault(c, []).append(r)
    out = []
    for c in sorted(by_col):
        rows = sorted(by_col[c])
        s = p = rows[0]
        for r in rows[1:] + [None]:
            if r is not None and r == p + 1:
                p = r
                continue
            out.append('%d %d %d %d' % (s, c, p, c))
            if r is not None:
                s = p = r
    return ','.join(out)


def write_book(path, inp, results, mode, with_test=True):
    """results: {(sheet, r, c): value} for the formula cells, 1-2 CH:AGS, ST_SUM, marks (1-2 R/S)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = {name: wb.create_sheet(name) for name in SHEETS}
    if not with_test:
        wb.remove(ws[S_TEST])
    # 1-1.Plan
    p = ws[S_PLAN]
    p['A1'] = 'Plan'
    p.cell(row=3, column=PLAN_COL_PN, value='P/N')
    for m in range(12):
        p.cell(row=3, column=PLAN_COL_M1 + m, value=MONTHS[m])
    for i, row in enumerate(inp['plan']):
        put(p, PLAN_ROW + i, PLAN_COL_PN, row[0])
        for m in range(12):
            put(p, PLAN_ROW + i, PLAN_COL_M1 + m, row[m + 1])
    # PART METHOD
    pm = ws[S_M]
    pm['A1'] = 'Parts made by another method  (read by CAPA_Recalc)'
    pm['A4'] = 'FMCNC1 machines = sheet 1-3 rows with Kind (column B):'
    pm['A5'] = 'HAND machines = sheet 1-3 rows with Kind (column B):'
    pm['A6'] = 'FMCNC2 machines = sheet 1-3 rows with Kind (column B):'
    for a, v in inp['ms'].items():
        pm[a] = v
    pm['A7'], pm['B7'], pm['C7'], pm['D7'] = 'Child P/N', 'Shop (FM / BD)', 'Method', 'Time'
    for i, (child, shop, how, t) in enumerate(inp['methods']):
        for j, v in enumerate([child, shop, how, t]):
            put(pm, 8 + i, 1 + j, v)
    # 1-2
    w = ws[S12]
    w['A1'] = 'ST table'
    w.cell(row=HDR_ROW + 1, column=COL_MK_FM, value='FM')
    w.cell(row=HDR_ROW + 1, column=COL_MK_BD, value='BD')
    w.cell(row=HDR_ROW + 1, column=COL_PN, value='ASSY P/N')
    for k in range(1, N_PROC + 1):
        w.cell(row=HDR_ROW, column=COL_ST + k - 1, value='Forming' if FM_P1 <= k <= FM_P2 else 'Group %d' % ((k - 1) // 10 + 1))
        w.cell(row=HDR_ROW + 1, column=COL_ST + k - 1, value=PROC_NAMES.get(k, 'Process %02d' % k))
    for m in range(12):
        c = COL_PLAN + m * BLOCK
        w.cell(row=HDR_ROW, column=c, value='Plan')
        w.cell(row=HDR_ROW + 1, column=c, value=MONTHS[m])
        for k in range(1, N_PROC + 1):
            put(w, COEF_ROW, c + k, inp['coef'][m][k - 1])
            w.cell(row=HDR_ROW + 1, column=c + k, value=PROC_NAMES.get(k, 'Process %02d' % k))
        w.cell(row=HDR_ROW + 1, column=COL_QTY + m, value="Q'TY " + MONTHS[m])
    for i, it in enumerate(inp['items']):
        r = FIRST_ROW + i
        put(w, r, COL_PN, it['pn'])
        put(w, r, COL_CHILD, it['child'])
        put(w, r, 9, 'part %d' % i)
        put(w, r, COL_PCS, it['pcs'])
        put(w, r, COL_OD, it['od'])
        put(w, r, COL_P, it['p'])
        for k, v in it['st'].items():
            put(w, r, COL_ST + k - 1, v)
    # ST_SUM headers (WriteSumHeaders)
    s = ws[S_SUM]
    s['A1'] = 'Totals of sheet 1-2, written by macro CAPA_Recalc - do not edit.'
    s.cell(row=5, column=COL_OD, value='OD')
    s.cell(row=5, column=COL_CH_FM, value='FM')
    s.cell(row=5, column=COL_CH_BD, value='BD')
    s.cell(row=5, column=COL_CHILD, value='ChildP/N')
    # 1-3
    t = ws[S13]
    t['H4'] = 'A/T'
    t.cell(row=HDR_ROW, column=CAPA_COL, value='A/T')
    for m in range(12):
        for k in range(1, N_PROC + 2):
            col = CAPA_COL + m * BLOCK + k - 1
            t.cell(row=3, column=col, value=float(20 + m % 3))
            t.cell(row=5, column=col, value='L1')
    for k in range(1, N_PROC + 1):
        t.cell(row=5, column=CAPA_FLAG + k - 1, value=PROC_NAMES.get(k, 'P%02d' % k))
    for i, mc in enumerate(inp['machines']):
        r = FIRST_ROW + i
        t.cell(row=r, column=1, value='L1')
        t.cell(row=r, column=2, value=mc['kind'])
        t.cell(row=r, column=3, value='MC-%02d' % i)
        t.cell(row=r, column=5, value=mc['e'])
        put(t, r, 7, mc['spec'])
        for k in mc['flags']:
            t.cell(row=r, column=CAPA_FLAG + k - 1, value=BULLET)
    # BY PROCESS
    b = ws[S_BP]
    b['B3'], b['C3'] = 'Process', 'Kind'
    for i, (k, kind) in enumerate(bp_layout()):
        b.cell(row=5 + i, column=2, value=PROC_NAMES.get(k, 'Process %02d' % k))
        if not kind.startswith('='):
            b.cell(row=5 + i, column=3, value=kind)
    # BY ITEM CODE
    ic = ws[S_IC]
    ic['G3'] = 'ChildP/N'
    for i, ch in enumerate(item_code_list(inp)):
        ic.cell(row=4 + i, column=7, value=ch)
    # TEST
    te = ws[S_TEST]
    for i, v in enumerate(TEST_INPUTS):
        put(te, 1 + i, 1, v)
    for i, v in enumerate(TEST_B):
        put(te, 1 + i, 2, v)
    for i, (a, v) in enumerate(TEST_TABLE):
        put(te, 1 + i, 5, a)
        put(te, 1 + i, 6, v)

    # results: 1-2 CH:AGS, marks, ST_SUM and (VALUES mode) the formula cells
    for (sh, r, c), v in results.items():
        if sh in (S12, S_SUM) and not is_formula_cell(inp, sh, r, c):
            put(ws[sh], r, c, v)
    fs = [f for f in formulas(inp) if with_test or f[0] != S_TEST]
    for sh, r, c, f in fs:
        if mode == 'FORMULAS':
            ws[sh].cell(row=r, column=c).value = r1c1_to_a1(f, r, c)
        else:
            put(ws[sh], r, c, results.get((sh, r, c)))
    # CAPA_RULES
    rs = ws[S_RULES]
    rs['A1'] = 'CAPA formula rules - written by macro CAPA_Setup. Do not edit.'
    rs['C1'] = mode
    rs['A2'], rs['B2'], rs['C2'], rs['D2'] = 'Sheet', 'Formula (R1C1)', 'Cells: row col row col, ...', 'Cells'
    rules = {}
    for sh, r, c, f in fs:
        rules.setdefault((sh, f), []).append((r, c))
    row = 3
    boxes = {}
    for (sh, f), cells in rules.items():
        rs.cell(row=row, column=1, value=sh)
        fc = rs.cell(row=row, column=2)
        fc.value = f
        fc.data_type = 's'                  # text, as the macro writes it ("'" & formula)
        rs.cell(row=row, column=3, value=rects_of(cells))
        rs.cell(row=row, column=4, value=len(cells))
        row += 1
        for r, c in cells:
            b0 = boxes.setdefault(sh, [0, 0])
            b0[0] = max(b0[0], r)
            b0[1] = max(b0[1], c)
    rs['F2'], rs['G2'], rs['H2'] = 'Sheet', 'Check name', 'Address'
    for i, (sh, (mr, mc)) in enumerate(sorted(boxes.items())):
        nm = 'CAPA_CHK_%d' % (i + 1)
        adr = '$A$1:$%s$%d' % (get_column_letter(mc), mr)
        rs.cell(row=3 + i, column=6, value=sh)
        rs.cell(row=3 + i, column=7, value=nm)
        rs.cell(row=3 + i, column=8, value=adr)
        dn = DefinedName(nm, attr_text="'%s'!%s" % (sh, adr), hidden=True)
        wb.defined_names[nm] = dn
    rs.sheet_state = 'veryHidden'
    wb.save(path)


_formula_cells = {}


def is_formula_cell(inp, sh, r, c):
    key = id(inp)
    if key not in _formula_cells:
        _formula_cells[key] = {(s, rr, cc) for s, rr, cc, _ in formulas(inp)}
    return (sh, r, c) in _formula_cells[key]


def oracle_results(inp):
    o = oracle(inp)
    res = {}
    for (r, c), v in o['region'].items():
        res[(S12, r, c)] = v
    for (r, c), v in o['stsum'].items():
        res[(S_SUM, r, c)] = v
    for i, (a, b) in enumerate(o['marks']):
        res[(S12, FIRST_ROW + i, COL_MK_FM)] = a
        res[(S12, FIRST_ROW + i, COL_MK_BD)] = b
    return res, o


def recalc_lo(src, dst):
    tmp = os.path.join(OUT, 'lo_out')
    shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(tmp)
    subprocess.run(['soffice', '--headless', '--norestore', '--convert-to', 'xlsx', '--outdir', tmp, src],
                   check=True, capture_output=True, timeout=600)
    shutil.move(os.path.join(tmp, os.path.basename(src)), dst)
    shutil.rmtree(tmp, ignore_errors=True)


def read_results(path, inp):
    """values of the formula cells of a calculated workbook"""
    wb = openpyxl.load_workbook(path, data_only=True)
    res = {}
    for sh, r, c, _ in formulas(inp):
        cell = wb[sh].cell(row=r, column=c)
        v = cell.value
        if cell.data_type == 'e':
            v = Err(v)
        elif isinstance(v, int) and not isinstance(v, bool):
            v = float(v)
        res[(sh, r, c)] = v
    return res


def make_shared(path, sheet_name):
    """Turns the formulas of each column of a sheet into shared formulas (as Excel writes them)."""
    zin = zipfile.ZipFile(path)
    wbx = zin.read('xl/workbook.xml').decode()
    rels = zin.read('xl/_rels/workbook.xml.rels').decode()
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(sheet_name), wbx).group(1)
    target = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
    if target is None:
        target = re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    part = 'xl/' + target.group(1).lstrip('/').replace('xl/', '')
    xml = zin.read(part).decode()
    cells = list(re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*)><f(?:\s+aca="[^"]*")?>([^<]*)</f>', xml))
    by_col = {}
    for m in cells:
        by_col.setdefault(m.group(1), []).append(m)
    repl = {}
    si = 0
    for col, ms in by_col.items():
        if len(ms) < 2:
            continue
        first, last = ms[0], ms[-1]
        ref = '%s%s:%s%s' % (col, first.group(2), col, last.group(2))
        repl[first.start()] = (first, '<c r="%s%s"%s><f t="shared" ref="%s" si="%d">%s</f>' % (
            col, first.group(2), first.group(3), ref, si, first.group(4)))
        for m in ms[1:]:
            repl[m.start()] = (m, '<c r="%s%s"%s><f t="shared" si="%d"/>' % (col, m.group(2), m.group(3), si))
        si += 1
    out, pos = [], 0
    for start in sorted(repl):
        m, text = repl[start]
        out.append(xml[pos:m.start()])
        out.append(text)
        pos = m.end()
    out.append(xml[pos:])
    new = ''.join(out)
    tmp = path + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new.encode() if item.filename == part else zin.read(item.filename)
            zout.writestr(item, data)
    zin.close()
    os.replace(tmp, path)
    return si


def jsonable(v):
    if isinstance(v, Err):
        return {'e': str(v)}
    return v


def main():
    os.makedirs(OUT, exist_ok=True)
    books = {}
    for variant in ['A', 'B']:
        inp = make_inputs(variant)
        ores, o = oracle_results(inp)
        raw = os.path.join(OUT, variant + '_formulas_raw.xlsx')
        write_book(raw, inp, ores, 'FORMULAS')
        calc = os.path.join(OUT, variant + '_formulas.xlsx')
        recalc_lo(raw, calc)
        os.remove(raw)
        res = dict(ores)
        res.update(read_results(calc, inp))
        books[variant] = (inp, res, o)
        print(variant, 'formula cells', len(formulas(inp)), 'groups', sum(1 for k in o['stsum'] if k[1] == COL_OD),
              'method rows not used', o['bad'])
    inp_a, res_a, _ = books['A']
    inp_b, res_b, _ = books['B']
    write_book(os.path.join(OUT, 'A_values.xlsx'), inp_a, res_a, 'VALUES')
    write_book(os.path.join(OUT, 'AB_values.xlsx'), inp_b, res_a, 'VALUES')
    # the sample of the web page: other inputs (without the error cases) and the old values of A
    inp_s = make_inputs('B', clean=True)
    sample = os.path.join(os.path.dirname(OUT), '..', 'src', 'ui', 'sample.xlsx')
    write_book(sample, inp_s, res_a, 'VALUES', with_test=False)
    n = make_shared(os.path.join(OUT, 'A_formulas.xlsx'), S_IC)
    print('shared formula groups in', S_IC, n)
    exp = [[sh, r, c, jsonable(v)] for (sh, r, c), v in sorted(res_b.items())]
    with open(os.path.join(OUT, 'expected_B.json'), 'w') as fh:
        json.dump(exp, fh)
    print('written to', OUT)


if __name__ == '__main__':
    sys.exit(main())
