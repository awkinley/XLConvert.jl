from functools import reduce
import math
import numpy as np
import pandas as pd
import datetime
from scipy import stats
from scipy import constants

import openpyxl
from openpyxl.utils.cell import range_boundaries

wb = openpyxl.load_workbook("Modular TEA - master - v1.25.xlsm", data_only=True)

def convert_xl_value(v):
    if v is None:
        return pd.NA
    if isinstance(v, datetime.date):
        return pd.Timestamp(v)
    if isinstance(v, str) and v.startswith("#"):
        return pd.NA
    return v

def load_range(sheet:str, first:str, last:str):
    ws = wb[sheet]
    min_col, min_row, max_col, max_row = range_boundaries(f"{first}:{last}")

    # Returns a 2D list of values (not Cell objects)
    return [
        list(map(convert_xl_value,row))
        for row in ws.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        )
    ]

class BlankZeroType(np.float64):
    """Represents a value that behaves mathematically as zero, but is tagged as being from a blank"""
    ...

BlankZero = BlankZeroType(0.0)


def not_none(a):
    return a is not None and not pd.isna(a)


def as_array(val):
    if isinstance(val, (pd.Series, pd.DataFrame)):
        return np.atleast_1d(val.values.squeeze())
    if isinstance(val, tuple):
        return val[0]

    return val

def flat(lst):
    for x in lst:
        if not isinstance(lst, str) and has_len(lst):
            arr = as_array(x)
            if has_len(arr):
                yield from flat(arr)
            else:
                yield arr
        else:
            yield x


def safe_len(a):
    try:
        if isinstance(a, str):
            return 0
        else:
            return len(a)
    except TypeError:
        return 0
    
def has_len(a):
    try:
        if isinstance(a, str):
            return False
        len(a)
        return True
    except TypeError:
        return False

def logical(a):
    if pd.isna(a):
        return False
    return bool(a)


def na_to_zero(a):
    if isinstance(a, np.ndarray):
        start_shape = a.shape
        return np.array(list(map(na_to_zero, a.flat)), dtype=a.dtype).reshape(start_shape)
    if isinstance(a, pd.Series):
        assert len(a) == 1
        a = a.iloc[0]

    if pd.isna(a) or (isinstance(a, float) and math.isnan(a)):
        return 0
    return a

def xlookup(val, ref, result, if_not_found=None, search_mode=0):
    # print(f"xl_xlookup({type(val)}, {type(ref)}, {type(result)})")
    if not isinstance(val, str) and safe_len(val) > 1:
        if isinstance(ref, tuple) and isinstance(result, tuple):
            return [xlookup(v, ref[0], result[0]) for v in as_array(val)]
        else:
            return [xlookup(v, ref[i], result[i]) for i, v in enumerate(as_array(val))]

    if isinstance(result, pd.DataFrame):
        # if isinstance(ref, pd.Series):
        #     print(f"{result.shape = }, {ref.shape = }, {len(ref) = }")
        if result.shape[0] > 1 and result.shape[1] > 1 and isinstance(ref, pd.DataFrame):
            result = result.values.T
            ref = as_array(ref)
        elif isinstance(ref, (pd.Series, list)) and len(ref) != result.shape[0]  and len(ref) == result.shape[1]:
            # print("treating ref as column-wise")
            result = result.values.T
            # print(f"{result.shape = }")
            ref = as_array(ref)
            # print(f"{ref.shape = }")
        else:
            ref = as_array(ref)
            result = as_array(result)
    else:
        ref = as_array(ref)
        result = as_array(result)

    if pd.isna(val):
        if if_not_found is not None:
            return if_not_found
        else:
            return val
    if result is pd.NA:
        if if_not_found is not None:
            return if_not_found
        else:
            return pd.NA
        
    # print("XLOOKUP ref")
    # print(ref)
    for test, res in zip(ref, result):
        if safe_len(test) > 1:
            test = test[0]
        if pd.isna(test):
            continue
        if test == val:
            if safe_len(res) == 0:
                if pd.isna(res):
                    return BlankZero
                else:
                    return res
                # return na_to_zero(res)
                # return res
            else:
                return res

    if search_mode == 1:
        min_test = None
        # print(f"{min_test = }")
        min_test_idx = None
        for i, test in enumerate(ref):
            # print(f"test = {test}, res = {result[i]}")
            if gt(test, val) and (min_test is None or test < min_test):
                # print(f"new lowest = {test}")
                min_test_idx = i
                min_test = test
        # print(f"{min_test_idx = }")
        if min_test_idx is None:
            return pd.NA
        res = result[min_test_idx]
        # print(f"{res = }")

        if safe_len(res) == 0:
            return na_to_zero(res)
        else:
            return res

    if search_mode == -1:
        min_test = None
        # print(f"{min_test = }")
        min_test_idx = None
        for i, test in enumerate(ref):
            # print(f"test = {test}, res = {result[i]}")
            if lt(test, val) and (min_test is None or test > min_test):
                # print(f"new lowest = {test}")
                min_test_idx = i
                min_test = test
        # print(f"{min_test_idx = }")
        if min_test_idx is None:
            return pd.NA
        res = result[min_test_idx]
        # print(f"{res = }")

        if safe_len(res) == 0:
            return na_to_zero(res)
        else:
            return res

    if if_not_found is not None:
        return if_not_found
    else:
        return pd.NA
    
def xmatch(val, ref, match_mode=None, search_mode=0):
    if not has_len(ref):
        ref = [ref]
    # print(f"{ref = }")
    # print(f"{len(as_array(ref)) = }")
    if isinstance(ref, pd.DataFrame):
        ref = as_array(ref)

    # print(f"{ref = }")
    # print(f"{safe_len(ref) = }")
    return xlookup(val, ref, range(safe_len(ref)), search_mode=search_mode)


def common_datetime_promote(a, b):
    if not (isinstance(a, datetime.date) or isinstance(b, datetime.date)):
        return a, b
    return promote_to_datetime(a), promote_to_datetime(b)
def eq(a, b):
    if pd.isna(a):
        return pd.isna(b) or b == 0 or b == ""
    if b == "":
        return a == b or pd.isna(a) or a is BlankZero # or a == 0
    return a == b

def gt(a, b):
    if any(map(pd.isna, (a, b))):
        return False

    if isinstance(a, bool) and is_number(b):
        return True

    if isinstance(b, bool) and is_number(a):
        return False

    if isinstance(a, str) and is_number(b):
        return True
    if isinstance(b, str) and is_number(a):
        return False

    return a > b

def leq(a, b):
    return lt(a, b) or not gt(a, b)
    # if any(map(pd.isna, (a, b))):
    #     return False
    # # print(f"{a = }, {b = }")
    # if isinstance(b, str) and is_number(a):
    #     return True
    # return a <= b

def geq(a, b):
    # return gt(a, b) or eq(a, b)
    return gt(a, b) or not lt(a, b)

def lt(a, b):
    if any(map(pd.isna, (a, b))):
        return True

    if isinstance(a, bool) and is_number(b):
        return False

    if isinstance(b, bool) and is_number(a):
        return True

    if isinstance(a, str) and is_number(b):
        return False
    if isinstance(b, str) and is_number(a):
        return True

    a, b = common_datetime_promote(a, b)
    return a < b

def as_num_array(a):
    a = as_array(a)
    return np.asarray(np.where(pd.isna(a), np.nan, a), dtype=float)

def sin(a):
    return np.sin(as_num_array(a))
def cos(a):
    return np.cos(as_num_array(a))
def asin(a):
    return np.arcsin(as_num_array(a))
def tan(a):
    return np.tan(as_num_array(a))
def atan(a):
    return np.arctan(as_num_array(a))
def radians(a):
    # print(f"{as_array(a) = }")
    # return np.deg2rad(as_array(a))
    return as_array(a) / 180 * np.pi

def degrees(a):
    # print(f"{as_array(a) = }")
    # return np.deg2rad(as_array(a))
    return as_array(a) * 180 / np.pi

def is_number(a):
    if isinstance(a, np.ndarray):
        return a.shape == ()
    return isinstance(a, (int, float, np.number))

def compare_list(a_list, b_list):
    for i, (a, b) in enumerate(zip(a_list, b_list)): 
        if not pd.isna(b):
            if not compare(a, b):
                print(f"compare failed at index {i = }")
                assert compare(a, b)
    return True

def compare(a, b):
    if pd.isna(b):
        res = True
        # res = pd.isna(a) or a == 0
    elif isinstance(a, str):
        # what a weird edge cast to handle the fact that excel and python spell
        # false with different capitalization
        if isinstance(b, str) and "FALSE" in b:
            a = a.replace("False", "FALSE")
        res = a == b
    elif pd.isna(a):
        res = pd.isna(b) or b == 0
    elif is_number(a) and math.isnan(a) :
        res = math.isnan(b) or b == 0
    elif is_number(a) and is_number(b):
        res = math.isclose(a, b)
        # res = abs(a - b) < 1e-10
    elif isinstance(a, (pd.Timestamp, np.datetime64)) and isinstance(b, (pd.Timestamp, np.datetime64)):
        res = abs(a - b) < pd.Timedelta(seconds=1)
    elif isinstance(a, datetime.datetime) and type(b) is datetime.date:
        res = a == promote_to_datetime(b)
    else:
        res = a == b
    if not res:
        print(f"{a} != {b}")
    return res


def xlsum(*args):
    all_args = flat(args)
    def make_numeric(a):
        if is_number(a):
            return a

        try:
            return float(a)
        except ValueError:
            return 0
    return sum(map(make_numeric, filter(not_none, all_args)))

def min(*args):
    # print(f"{args = }")
    all_args = list(flat(args))
    # print(f"{all_args = }")
    if any(map(pd.isna, all_args)):
        return pd.NA

    # print(f"{all_args = }")
    if len(all_args) == 0:
        return 0

    all_args = list(filter(is_number, all_args))

    return np.min(all_args)

def max(*args):
    # print(f"{args = }")
    all_args = list(flat(args))
    # print(f"{all_args = }")
    if any(map(pd.isna, all_args)):
        return np.max(list(filter(not_none, all_args)))
        # return pd.NA
    if len(all_args) == 0:
        return 0
    return np.max(all_args)

def average(*args):
    if len(args) == 1 and isinstance(args[0], pd.Series):
        series = args[0].infer_objects()
        if series.dtype == "datetime64[ns]":
            # Kind of a silly way to average dates, but it works because
            # timedelta's can be averaged, unlike datetimes
            t0 = series.iloc[0] 
            return np.mean(series - t0) + t0
    if len(args) == 2:
        if all(isinstance(a, pd.Timestamp) for a in args):
            a, b = args
            return a + (b - a) / 2

    all_args = as_array(list(filter(not_none, flat(args))))
    return np.mean(all_args)


def switch(val, *args):
    if safe_len(val) > 1:
        return [switch(v, *args) for v in val]

    # print(f"{val = }")
    # print(*args)

    assert len(args) % 2 == 0
    n = len(args) // 2

    for i in range(n):
        if eq(val, args[2 * i]):
            return args[2 * i + 1]
    return pd.NA
    raise ValueError(f"xl_switch couldn't find looked up value {val = }, {args = }")

def mod(a, b):
    return a % b

def linest(ys, xs, const=True, stats=False):
    if np.any(pd.isna(xs)):
        return pd.NA

    xs = np.asarray(xs, dtype=np.float64)
    ys = np.asarray(ys, dtype=np.float64)

    if const:
        xs = np.hstack([xs, np.ones((xs.shape[0], 1))])
    
    # print("xs", xs)
    # print("ys", ys)
    res = np.linalg.lstsq(xs, np.squeeze(ys))[0]
    # print(f"{res = }")
    return res[-2]

    # print(ys)
    # print(xs)
    # return pd.NA

def isnumber(a):
    if isinstance(a, np.ndarray) and a.shape == ():
        return True
    return isinstance(a, (int, float, np.floating, np.integer))

def find(pattern, string, start_num=1):
    if isinstance(string, str):
        res = string.find(pattern)
        if res == -1:
            return pd.NA
        else:
            return res
    return pd.NA
    
def na_to_false(v):
    if pd.isna(v):
        return False
    return v

def xlall(vals):
    return all(map(na_to_false, vals))

def xlany(vals):
    return any(map(na_to_false, vals))

def is_whole_number(a):
    if not is_number(a):
        return False
    
    return a == int(a)

def promote_to_datetime(date):
    if type(date) is datetime.date:
        return datetime.datetime(date.year, date.month, date.day)
    return date
    
def date_add(date, delta):
    # print(f"{date = }, {delta = }")
    # if type(date) is datetime.date:
    #     if is_whole_number(delta):
    #         return date + datetime.timedelta(days=delta)
    #     else:
    #         return promote_to_datetime(date) + datetime.timedelta(days=delta)
    # else:
    #     # print(f"{date + delta = }")
    #     return date + datetime.timedelta(delta)
    # print(f"{delta = }")
    # pandas is sometimes less accurate than datetime, so we round to microseconds
    if isinstance(delta, np.ndarray):
        time_delta = delta.astype(float) * pd.Timedelta(1, "D")
        date = date.astype("datetime64")
        # print("date = ")
        # print(date)
        # print(f"{type(date) = }")
        # print(f"{date.dtype = }")
        res =  pd.DatetimeIndex(date + time_delta).round(freq="us").values
    else:
        if isinstance(delta, np.datetime64):
            zero_day = pd.Timestamp(1900, 1, 1) - pd.Timedelta(days=2)
            delta = date_sub(delta, zero_day)

        time_delta = pd.Timedelta(delta, "D")
        # print(date)
        res =  (date + time_delta).round(freq="us")
    return res





def date_sub(date, delta):
    # print(f"{date = }, {delta = }")
    if date is False:
        # This is what excel interprets a day of 0 to be
        date = datetime.date(1900, 1, 1) - datetime.timedelta(days=2)
    if isinstance(date, datetime.date) and isinstance(delta, datetime.date):
        date = promote_to_datetime(date)
        delta = promote_to_datetime(delta)
        return (date - delta) / datetime.timedelta(days=1)

    if type(date) is datetime.date:
        if is_whole_number(delta):
            return date - datetime.timedelta(days=delta)
        else:
            return promote_to_datetime(date) - datetime.timedelta(days=delta)
    if isinstance(delta, pd.Timestamp):
        # print(f"{date = }")
        # print(f"{delta = }")
        # print(f"{(date - delta) = }")
        # print(f"{(date - delta) / pd.Timedelta(days=1.0) = }")
        return (date - delta) / pd.Timedelta(days=1.0) 

    if isinstance(delta, np.ndarray):
        return (date - delta) / pd.Timedelta(days=1.0)

    # print(f"{delta = }")
    return date - datetime.timedelta(days=delta)

def year(a):
    if not_none(a):
        return pd.Timestamp(a).year
    return pd.NA

def month(a):
    if not_none(a):
        return pd.Timestamp(a).month
    return pd.NA

def day(a):
    return pd.Timestamp(a).day
    # return a.day

def days(end, start):
    return (end - start).days

def concat(a, b):
    def try_int(a):
        if is_whole_number(a):
            return int(a)
        return a

    # print(f"{a = }, {b = }")
    if isinstance(a, pd.Series):
        na_mask = pd.isna(a)
        a = a.astype(str)
        a[na_mask] = pd.NA
    if isinstance(b, pd.Series):
        na_mask = pd.isna(b)
        b = b.astype(str)
        b[na_mask] = pd.NA
    
    if isinstance(a, pd.Series) and isinstance(b, pd.Series):
        na_mask = pd.isna(a) | pd.isna(b)
        a_new = a.apply(try_int)
        b_new = b.apply(try_int)
        # print("b = ")
        # print(b)
        # print("b_new = ")
        # print(b_new)
        res = a_new.astype(str) + b_new.astype(str)
        res[na_mask] = ""
        return res
    
    if is_number(a):
        if is_whole_number(a):
            a = int(a)
        a = str(a)
    if is_number(b):
        if is_whole_number(b):
            b = int(b)
        b = str(b)

    # print(f"{a = }, {b = }")
    return a + b

def xlround(a, digits=0):
    if isinstance(a, pd.Timestamp) and digits==0:
        return a.round(freq="D")
    if is_whole_number(digits):
        digits = int(digits)
    return round(a, digits)

def roundup(a, digits=0):
    assert digits == 0
    return np.ceil(a)
    # if isinstance(a, pd.Timestamp) and digits==0:
    #     return a.round(freq="D")
    # if is_whole_number(digits):
    #     digits = int(digits)
    # return round(a, digits)

def rounddown(a, digits=0):
    assert digits == 0
    return np.floor(a)

def floor(a, significance):
    return np.floor(a / significance) * significance

def iferror(a, default):
    if a is BlankZero:
        return 0

    if not_none(a):
        return a
    
    return default

def sum_product(*args):
    # print("sum_product - Argument lengths")
    # print(list(map(safe_len, args)))
    def mult(a, b):
        try:
            return as_array(a) * as_array(b)
        except ValueError as e:
            print("="*10, "a", "="*10)
            print(a)
            print("="*10, "b", "="*10)
            print(b)
            print(b.shape)
            raise e


    return np.float64(xlsum(reduce(mult, args)))


def match(value, search_range, match_mode=0):
    if not has_len(search_range):
        search_range = [search_range]
    for i, val in enumerate(search_range):
        if eq(value, val):
            return i
    return pd.NA

def istext(a):
    return isinstance(a, str)

def match_case_insensitive(a, b):
    # print("==== a value ====")
    # print(a)
    # print("==== b value ====")
    # print(b)
    # flattened_a = False

    if isinstance(a, pd.DataFrame):
        if a.shape[0] == 1:
            a = a.iloc[0, :]
            # flattened_a = True
        # assert a.shape[1] == 1, f"{a.shape = }"
        # a = a.iloc[:, 0]
    if not isinstance(b, str):
        # print(f"{a = }, {b = }")
        # print(list(a))
        # print(type(b))
        if isinstance(b, pd.Series):
            assert len(b) == 1
            b = b.iloc[0]
        mask =  np.array([eq(v, b) for v in a])
        # print(mask)
        return mask

    return (a.astype(str).str.lower() == b.lower()).values

def convert(val, in_unit, out_unit):
    if in_unit == "HP" and out_unit == "W":
        return val * constants.hp
    if in_unit == "kn":
        mtr_per_sec = val * constants.knot
        # meters per hour
        if out_unit == "m/hr":
            return mtr_per_sec * 3600
    if in_unit == "ha" and out_unit == "us_acre":
        # print(f"{val = }")
        us_acre = 4046 + (13525426/15499969)
        return val * (constants.hectare / us_acre)
    if in_unit == "ft^3" and out_unit == "m^3":
        return val * (constants.foot ** 3)

    raise ValueError(f"Unknown unit conversion {in_unit} => {out_unit}")

def norm_dist(x, mu, sigma, is_cumulative):
    if is_cumulative:
        return stats.norm.cdf(x, mu, sigma)
    else:
        return stats.norm.pdf(x, mu, sigma)

def pmt(rate, nper, pv):
    # @info "xl_pmt" rate nper pv
    return -1 * np.sign(pv) * (pv * rate) / (1 - (1 + rate)**(-nper))

def safe_recip(val):
    if val == 0:
        return 0

    return 1 / val