//! The date functions (KERNEL.8, slice d, 2026-10-20): `DATE`, `YEAR`,
//! `MONTH`, `DAY`, `WEEKDAY`, `EOMONTH`, `EDATE`, `DAYS`, `NETWORKDAYS`,
//! `WORKDAY`, `TIME`, `HOUR`, `MINUTE`, `SECOND`, `DATEDIF`, `YEARFRAC` and
//! `WEEKNUM`. The language's own, since a game reaches the language alone
//! (Alonzo's `AD-1`) and a game's calendar, Buffet's sessions and the
//! holidays its Caveats promise, is a date's work; Excel's library has them
//! through it (the owner's decision, 2026-10-20).
//!
//! A date is a serial, a count of days in Excel's 1900 calendar, which is
//! here ([`civil`], [`serial_of`]) and which `numfmt` draws a date's code
//! by: 1 is 1 January 1900, 60 the 29 February 1900 that never was,
//! Lotus's, which Excel keeps, 61 is 1 March 1900, 0 the day before the
//! first, `1900-01-00`, and 2,958,465 is 31 December 9999 ([`LAST_DAY`]);
//! a time is the fraction of a day. Serial 1 is a Sunday, which the true 1
//! January 1900 was not: Excel counts weekdays by its serials. A serial
//! below 0 or past the last day is `#NUM!`. Each rule is Excel's as its
//! documentation states it, held where the documentation is silent to
//! Excel's own saved answer (`scripts/recalc/dates.txt`).
//!
//! A value is read as arithmetic reads it (`eval::numeric`): a number, a
//! truth value 1 or 0, an empty value 0, a text spelling a number its
//! number; a text holding a digit that no rule here reads, as `"1/1/2026"`
//! is, labelled (`a locale's number`), since a host reads it by its locale,
//! and any other text `#VALUE!` (the owner's decision, 2026-10-20). A whole
//! number an argument asks for is read as Excel's save answers it: `DATE`'s
//! and `TIME`'s parts, `WORKDAY`'s days, a type and a basis by their floor,
//! `EOMONTH`'s and `EDATE`'s months truncated toward zero. A date's parts,
//! its weekday among them, and a time's are read from the serial rounded
//! to the second, as `numfmt` draws them, so that a time a hair before
//! midnight is the next day's; a date the other functions count from is
//! its serial's whole part.
//!
//! Two edges are labelled, where Excel's saves of one formula answered
//! differently, so that no answer is Excel's: a year outside the calendar
//! that `DATE`'s months carry back into it, and `WEEKNUM`'s types past 21,
//! which its documentation does not name.
//!
//! What a locale or a clock decides is not here: `DATEVALUE` and
//! `TIMEVALUE`, a date or a time spelled in a text, and `NOW` and `TODAY`,
//! whose answer changes with no cell changing, each named not computed by a
//! door and refused by an engine's load. A workbook counting its dates
//! from 1904, as a Mac's may (`date1904`), has every function here whose
//! answer turns on the calendar labelled by a door ([`label_1904`]).

use super::eval::{numeric, Arg, Ctx, Library, Reason};
use super::texts::{answer, wrong, Fault};
use super::value::{ErrorKind, Value};
use crate::numfmt::Code;

/// The last day Excel's calendar holds, 31 December 9999.
pub(crate) const LAST_DAY: i64 = 2_958_465;

/// The proleptic Gregorian date `z` days after 1 January 1970 (Howard
/// Hinnant's `civil_from_days`).
fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let z = z + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z.rem_euclid(146_097);
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (yoe + era * 400 + i64::from(m <= 2), m, d)
}

/// The days from 1 January 1970 to a proleptic Gregorian date, its month
/// 1 to 12 (Howard Hinnant's `days_from_civil`).
fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = y.div_euclid(400);
    let yoe = y.rem_euclid(400);
    let doy = (153 * ((m + 9) % 12) + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

/// The year, the month and the day of a serial in Excel's 1900 calendar: 0
/// the day before the first, `1900-01-00`, and 60 the 29 February 1900 that
/// never was; every other serial the proleptic Gregorian calendar's day,
/// 25,569 being 1 January 1970 and a serial below 60 one day later than its
/// count, the day that never was not among them. Below 0 the days before
/// the calendar, which a week's number reaches.
pub(crate) fn civil(serial: i64) -> (i64, i64, i64) {
    match serial {
        0 => (1900, 1, 0),
        60 => (1900, 2, 29),
        s if s < 60 => civil_from_days(s - 25_568),
        s => civil_from_days(s - 25_569),
    }
}

/// The serial of a month's first day, its month 1 to 12: January and
/// February 1900 counted from the first day, every later month a day
/// further on for the day that never was.
fn first_of_month(year: i64, month: i64) -> i64 {
    let days = days_from_civil(year, month, 1);
    if (year, month) < (1900, 3) {
        days + 25_568
    } else {
        days + 25_569
    }
}

/// A month counted from January of year 0, as a year and a month.
fn month_of(count: i64) -> (i64, i64) {
    (count.div_euclid(12), count.rem_euclid(12) + 1)
}

/// The serial of a year, a month and a day as `DATE` carries them: the
/// months past twelve or below one into the years, the days past the
/// month's end or below its first into the months, by counting from the
/// month's first day. February 1900 holds its 29th.
pub(crate) fn serial_of(year: i64, month: i64, day: i64) -> i64 {
    let (y, m) = month_of(year * 12 + month - 1);
    first_of_month(y, m) + day - 1
}

/// How many days a month of the proleptic Gregorian calendar holds,
/// February 1900 28: the month's length `EOMONTH`, `EDATE` and `YEARFRAC`
/// count by, where the serials hold the 29th that never was (Excel's
/// save, `EOMONTH(60,0)` 59).
fn days_in(year: i64, month: i64) -> i64 {
    let (ny, nm) = month_of(year * 12 + month);
    days_from_civil(ny, nm, 1) - days_from_civil(year, month, 1)
}

/// A month the calendar holds: a month counted from January of year 0, as
/// a year and a month, its year 1900 to 9999, else `#NUM!`. `DATE`,
/// `EOMONTH` and `EDATE` hold the month they reach to it, before its day
/// is counted, so that `DATE(9999,13,-31)` and `DATE(1900,0,31)` are
/// `#NUM!` and `EDATE(0,0)` is 0 (Excel's saves).
fn held_month(count: i64) -> Result<(i64, i64), Fault> {
    let (y, m) = month_of(count);
    if !(1900..=9999).contains(&y) {
        return num();
    }
    Ok((y, m))
}

/// A serial's day of the week, 0 a Sunday to 6 a Saturday, serial 1 a
/// Sunday.
fn weekday(serial: i64) -> i64 {
    (serial + 6).rem_euclid(7)
}

/// The answer `#NUM!`: a date or a time Excel's calendar does not hold, or
/// an argument out of its range.
fn num<T>() -> Result<T, Fault> {
    Err(Fault::Error(ErrorKind::Num))
}

/// An argument read as arithmetic reads it (see the module's comment).
fn number(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Result<f64, Fault> {
    Ok(numeric(&ctx.one(arg)?)??)
}

/// A date argument's day: its serial's whole part, `#NUM!` below 0 or
/// past the calendar's last day.
fn day(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Result<i64, Fault> {
    let x = number(ctx, arg)?;
    if !(0.0..(LAST_DAY + 1) as f64).contains(&x) {
        return num();
    }
    Ok(x.floor() as i64)
}

/// A date and time argument as `numfmt` draws one: its serial rounded to
/// the second, as its day and its second of that day, so that a time a
/// hair before midnight is the next day's midnight; `#NUM!` below 0, or
/// where the day rounded to is past the calendar's last (Excel's save,
/// `YEAR(2958465.9999999)`).
fn moment(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Result<(i64, i64), Fault> {
    let x = number(ctx, arg)?;
    let seconds = (x * 86_400.0).round();
    if x < 0.0 || seconds >= ((LAST_DAY + 1) * 86_400) as f64 {
        return num();
    }
    let seconds = seconds as i64;
    Ok((seconds / 86_400, seconds % 86_400))
}

/// The span every whole number an argument asks for is held to, wider than
/// the calendar, so that no sum of them overflows.
const SPAN: f64 = 1e9;

/// A whole number `DATE`, `TIME` and `WORKDAY` ask for, and a type or a
/// basis: its number's floor, as Excel's save answers (`DATE(2026,1,-0.5)`
/// the year's 30 December, `WORKDAY` of -1.9 days two, `YEARFRAC`'s basis
/// of -0.5 `#NUM!`).
fn floored(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Result<i64, Fault> {
    Ok(number(ctx, arg)?.floor().clamp(-SPAN, SPAN) as i64)
}

/// A whole number `EOMONTH` and `EDATE` ask for, their months: its number
/// truncated toward zero, as Excel documents them and its save answers
/// (`EOMONTH` of -1.9 months one).
fn truncated(ctx: &Ctx<'_>, arg: &Arg<'_>) -> Result<i64, Fault> {
    Ok(number(ctx, arg)?.trunc().clamp(-SPAN, SPAN) as i64)
}

/// An answer labelled, `what` its reason: an edge where Excel's saves of
/// one formula answer differently, so that no answer is Excel's.
fn label<T>(what: &str) -> Result<T, Fault> {
    Err(Fault::Reason(Reason::Construct(what.to_string())))
}

/// A serial as an answer, `#NUM!` outside the calendar.
fn dated(serial: i64) -> Result<Value, Fault> {
    if (0..=LAST_DAY).contains(&serial) {
        Ok(Value::Number(serial as f64))
    } else {
        num()
    }
}

/// `DATE(year, month, day)`: the serial of the date, each part its floor, a
/// year 0 to 1899 counted from 1900, as Excel counts it; the months past
/// twelve or below one carried into the years, the year they reach held to
/// the calendar ([`held_month`]), then the days past the month's end or
/// below its first carried into the months; `#NUM!` for a date the
/// calendar does not hold. A year below 0 or past 9999 is `#NUM!`, as Excel
/// documents it, but labelled where its months carry it into 0 to 9999:
/// Excel's saves answered `DATE(10000,-11,1)` 2958101 and then `#NUM!`.
pub(crate) fn date(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [y, m, d] = args else { return wrong() };
        let (y, m, d) = (floored(ctx, y)?, floored(ctx, m)?, floored(ctx, d)?);
        if !(0..=9999).contains(&y) {
            return if (0..=9999).contains(&(y * 12 + m - 1).div_euclid(12)) {
                label("a year outside the calendar")
            } else {
                num()
            };
        }
        let year = if y < 1900 { y + 1900 } else { y };
        let (y, m) = held_month(year * 12 + m - 1)?;
        dated(first_of_month(y, m) + d - 1)
    })
}

/// `YEAR(serial)`, `MONTH(serial)` and `DAY(serial)`: a part of its date,
/// the serial rounded to the second ([`moment`]).
fn part(ctx: &Ctx<'_>, args: &[Arg], pick: fn((i64, i64, i64)) -> i64) -> Result<Value, Reason> {
    answer(|| {
        let [s] = args else { return wrong() };
        Ok(Value::Number(pick(civil(moment(ctx, s)?.0)) as f64))
    })
}

/// `YEAR(serial)`.
pub(crate) fn year(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    part(ctx, args, |(y, _, _)| y)
}

/// `MONTH(serial)`.
pub(crate) fn month(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    part(ctx, args, |(_, m, _)| m)
}

/// `DAY(serial)`.
pub(crate) fn day_(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    part(ctx, args, |(_, _, d)| d)
}

/// The day a week of `type` begins on, 0 a Sunday, for `WEEKDAY`'s types
/// 11 to 17 and `WEEKNUM`'s: 11 a Monday to 16 a Saturday, 17 a Sunday.
fn week_start(kind: i64) -> i64 {
    (kind - 10) % 7
}

/// `WEEKDAY(serial, [type])`: its day of the week, the serial rounded to
/// the second as a date's parts are ([`moment`]), by type 1 (the default)
/// 1 a Sunday to 7 a Saturday, 2 1 a Monday to 7 a Sunday, 3 0 a Monday to
/// 6 a Sunday, 11 to 17 1 the day the type names, 11 a Monday to 17 a
/// Sunday; any other type `#NUM!`.
pub(crate) fn weekday_(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(1..=2).contains(&args.len()) {
            return wrong();
        }
        let d = weekday(moment(ctx, &args[0])?.0);
        let kind = match args.get(1) {
            None => 1,
            Some(a) => floored(ctx, a)?,
        };
        let w = match kind {
            1 => d + 1,
            2 => (d + 6) % 7 + 1,
            3 => (d + 6) % 7,
            11..=17 => (d - week_start(kind)).rem_euclid(7) + 1,
            _ => return num(),
        };
        Ok(Value::Number(w as f64))
    })
}

/// A date's month counted from January of year 0, and its day.
fn month_count(serial: i64) -> (i64, i64) {
    let (y, m, d) = civil(serial);
    (y * 12 + m - 1, d)
}

/// `EOMONTH(start, months)`: the serial of the last day of the month
/// `months` from the start's, February 1900's the 28th ([`days_in`]); a
/// month the calendar does not hold `#NUM!` ([`held_month`]).
pub(crate) fn eomonth(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [s, k] = args else { return wrong() };
        let (count, _) = month_count(day(ctx, s)?);
        let (y, m) = held_month(count + truncated(ctx, k)?)?;
        dated(first_of_month(y, m) + days_in(y, m) - 1)
    })
}

/// `EDATE(start, months)`: the serial of the same day `months` from the
/// start, the month's last where it has no such day, February 1900's the
/// 28th ([`days_in`]); a month the calendar does not hold `#NUM!`
/// ([`held_month`]). The day before the first, day 0, moves as January's
/// day 0, so that `EDATE(0,0)` is 0 (Excel's save).
pub(crate) fn edate(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [s, k] = args else { return wrong() };
        let (count, d) = month_count(day(ctx, s)?);
        let (y, m) = held_month(count + truncated(ctx, k)?)?;
        dated(first_of_month(y, m) + d.min(days_in(y, m)) - 1)
    })
}

/// `DAYS(end, start)`: the days from the start's date to the end's.
pub(crate) fn days(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [e, s] = args else { return wrong() };
        let (e, s) = (day(ctx, e)?, day(ctx, s)?);
        Ok(Value::Number((e - s) as f64))
    })
}

/// A function's holidays: each date the argument holds, a reference's
/// cells passed by where empty, its first error the answer, a text a
/// number as arithmetic reads one and a truth value `#VALUE!`; the
/// weekdays among them, each once, in order, so that a holiday on a
/// weekend or given twice counts once or not at all.
fn holidays(ctx: &Ctx<'_>, arg: Option<&Arg<'_>>) -> Result<Vec<i64>, Fault> {
    let mut values = Vec::new();
    match arg {
        None => {}
        Some(Arg::Area { areas, .. }) => ctx.each(areas, |v| {
            values.push(v);
            Ok(())
        })?,
        Some(a) => values.push(ctx.one(a)?),
    }
    if let Some(k) = values.iter().find_map(|v| match v {
        Value::Error(k) => Some(*k),
        _ => None,
    }) {
        return Err(Fault::Error(k));
    }
    let mut out = Vec::with_capacity(values.len());
    for v in &values {
        if let Value::Bool(_) = v {
            return wrong();
        }
        let x = numeric(v)??;
        if !(0.0..(LAST_DAY + 1) as f64).contains(&x) {
            return num();
        }
        let d = x.floor() as i64;
        if !matches!(weekday(d), 0 | 6) {
            out.push(d);
        }
    }
    out.sort_unstable();
    out.dedup();
    Ok(out)
}

/// The weekdays from `a` to `b`, both counted, `a` not after `b`.
fn weekdays_between(a: i64, b: i64) -> i64 {
    let days = b - a + 1;
    let mut n = days / 7 * 5;
    for s in a + days / 7 * 7..=b {
        if !matches!(weekday(s), 0 | 6) {
            n += 1;
        }
    }
    n
}

/// `NETWORKDAYS(start, end, [holidays])`: the weekdays from the start's
/// date to the end's, both counted, but the holidays among them; below 0
/// where the end comes first.
pub(crate) fn networkdays(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(2..=3).contains(&args.len()) {
            return wrong();
        }
        let (a, b) = (day(ctx, &args[0])?, day(ctx, &args[1])?);
        let holidays = holidays(ctx, args.get(2))?;
        let (lo, hi) = (a.min(b), a.max(b));
        let off = holidays.iter().filter(|&&h| lo <= h && h <= hi).count() as i64;
        let n = weekdays_between(lo, hi) - off;
        Ok(Value::Number(if a <= b { n } else { -n } as f64))
    })
}

/// `WORKDAY(start, days, [holidays])`: the date `days` weekdays from the
/// start's, the days their floor, before it for a count below 0, a holiday
/// passed by as a weekend is; the start's date itself for 0; `#NUM!`
/// outside the calendar.
pub(crate) fn workday(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(2..=3).contains(&args.len()) {
            return wrong();
        }
        let start = day(ctx, &args[0])?;
        let mut left = floored(ctx, &args[1])?;
        let holidays = holidays(ctx, args.get(2))?;
        let step = left.signum();
        let mut d = start;
        while left != 0 {
            d += step;
            if !(0..=LAST_DAY).contains(&d) {
                return num();
            }
            if !matches!(weekday(d), 0 | 6) && holidays.binary_search(&d).is_err() {
                left -= step;
            }
        }
        dated(d)
    })
}

/// `TIME(hour, minute, second)`: the fraction of a day the three make, each
/// its floor, past a day the fraction alone, taken from the days the
/// seconds make as Excel's save takes it, to the bit (`TIME(25,0,0)`
/// 0.04166666666666674, not 1/24); `#NUM!` for a part past 32,767 or a time
/// below 0.
pub(crate) fn time(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [h, m, s] = args else { return wrong() };
        let (h, m, s) = (floored(ctx, h)?, floored(ctx, m)?, floored(ctx, s)?);
        if h > 32_767 || m > 32_767 || s > 32_767 {
            return num();
        }
        let seconds = h * 3600 + m * 60 + s;
        if seconds < 0 {
            return num();
        }
        let days = seconds as f64 / 86_400.0;
        Ok(Value::Number(days - days.floor()))
    })
}

/// `HOUR(serial)`, `MINUTE(serial)` and `SECOND(serial)`: a part of its
/// time of day, the serial rounded to the second ([`moment`]), midnight
/// past the day's last second.
fn clock(ctx: &Ctx<'_>, args: &[Arg], pick: fn(i64) -> i64) -> Result<Value, Reason> {
    answer(|| {
        let [s] = args else { return wrong() };
        Ok(Value::Number(pick(moment(ctx, s)?.1) as f64))
    })
}

/// `HOUR(serial)`.
pub(crate) fn hour(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    clock(ctx, args, |s| s / 3600)
}

/// `MINUTE(serial)`.
pub(crate) fn minute(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    clock(ctx, args, |s| s / 60 % 60)
}

/// `SECOND(serial)`.
pub(crate) fn second(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    clock(ctx, args, |s| s % 60)
}

/// `DATEDIF(start, end, unit)`: the whole years (`"Y"`), months (`"M"`) or
/// days (`"D"`) from the start's date to the end's; the months past the
/// years (`"YM"`); the days past the months (`"MD"`), the end against the
/// start's day in the end's month or the one before, as Excel counts them,
/// which may fall below 0; and the days past the years (`"YD"`), the days
/// of the months past the years walked from the start's own month in its
/// own year, and then the days past the months, as Excel reckons them, the
/// one reading of its three saves' fourteen answers (22 June 1988 to 11 May
/// 2012 323 days, the walk passing February 1989, not 2012's 29 days; 29
/// February 2020 to 28 February 2022 365, February 2020's). A unit in
/// either case; any other unit, and an end before the start, `#NUM!`.
pub(crate) fn datedif(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        let [s, e, unit] = args else { return wrong() };
        let (s, e) = (day(ctx, s)?, day(ctx, e)?);
        let unit = ctx.one(unit)?.text()?.to_ascii_uppercase();
        if s > e {
            return num();
        }
        let (y1, m1, d1) = civil(s);
        let (y2, m2, d2) = civil(e);
        let months = (y2 - y1) * 12 + m2 - m1 - i64::from(d2 < d1);
        let md = e - serial_of(y2, m2 - i64::from(d2 < d1), d1);
        let n = match unit.as_str() {
            "D" => e - s,
            "M" => months,
            "Y" => months / 12,
            "YM" => months % 12,
            "MD" => md,
            "YD" => {
                let (y, m) = month_of(y1 * 12 + m1 - 1 + months % 12);
                first_of_month(y, m) - first_of_month(y1, m1) + md
            }
            _ => return num(),
        };
        Ok(Value::Number(n as f64))
    })
}

/// Whether a year of the proleptic Gregorian calendar holds a 29 February,
/// 1900 not among them: `YEARFRAC` counts a year's length so, though its
/// days are the serials' (Excel's save: 1 January 1900 to 1 January 1901
/// 366 days over a year of 365).
fn leap(year: i64) -> bool {
    (year % 4 == 0 && year % 100 != 0) || year % 400 == 0
}

/// Whether a date is its month's last day ([`days_in`]): 28 February 1900
/// is, and serial 60, its 29th, is not (Excel's save).
fn month_end(y: i64, m: i64, d: i64) -> bool {
    d == days_in(y, m)
}

/// `YEARFRAC(start, end, [basis])`: the years from one date to the other,
/// in either order, counted by the basis: 0 (the default) 30 days a month
/// and 360 a year by the US rule, the month's last day of February taken
/// for its 30th; 1 the days as they are over the length of the years they
/// span, a year of 366 days where a 29 February falls within a year's
/// span, else the mean of the years' lengths ([`leap`]); 2 the days over
/// 360; 3 the days over 365; 4 30 days a month by the European rule. Any
/// other basis `#NUM!`. The rules are Excel's as David Wheeler's reading
/// of it states them, held to Excel's save.
pub(crate) fn yearfrac(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(2..=3).contains(&args.len()) {
            return wrong();
        }
        let (a, b) = (day(ctx, &args[0])?, day(ctx, &args[1])?);
        let basis = match args.get(2) {
            None => 0,
            Some(x) => floored(ctx, x)?,
        };
        if !(0..=4).contains(&basis) {
            return num();
        }
        let (s, e) = (a.min(b), a.max(b));
        Ok(Value::Number(if s == e {
            0.0
        } else {
            year_fraction(s, e, basis)
        }))
    })
}

/// [`yearfrac`]'s count between two dates, the first before the second.
fn year_fraction(s: i64, e: i64, basis: i64) -> f64 {
    let (y1, m1, d1) = civil(s);
    let (y2, m2, d2) = civil(e);
    let thirty = |d1: i64, d2: i64| ((y2 - y1) * 360 + (m2 - m1) * 30 + d2 - d1) as f64 / 360.0;
    match basis {
        0 => {
            let (last1, last2) = (
                m1 == 2 && month_end(y1, m1, d1),
                m2 == 2 && month_end(y2, m2, d2),
            );
            // Wheeler's order of the rule's five cases, the three that make
            // both days 30 disjoint from the two that make the first 30.
            let (d1, d2) = if (d2 == 31 && d1 >= 30) || (last1 && last2) {
                (30, 30)
            } else if d1 == 31 || last1 {
                (30, d2)
            } else {
                (d1, d2)
            };
            thirty(d1, d2)
        }
        1 => {
            let days = (e - s) as f64;
            let within_a_year = y1 == y2 || (y2 == y1 + 1 && (m1, d1) >= (m2, d2));
            if within_a_year {
                let feb29 = (y1 == y2 && leap(y1))
                    || (leap(y1) && (m1, d1) <= (2, 29))
                    || (leap(y2) && ((m2, d2) >= (3, 1) || (m2, d2) == (2, 29)));
                days / if feb29 { 366.0 } else { 365.0 }
            } else {
                let years = (y2 - y1 + 1) as f64;
                let span = (days_from_civil(y2 + 1, 1, 1) - days_from_civil(y1, 1, 1)) as f64;
                days / (span / years)
            }
        }
        2 => (e - s) as f64 / 360.0,
        3 => (e - s) as f64 / 365.0,
        _ => thirty(d1.min(30), d2.min(30)),
    }
}

/// The ISO week of a date: its week's Thursday's, counted from the week
/// holding that year's first Thursday.
fn iso_week(serial: i64) -> i64 {
    let thursday = serial - (weekday(serial) + 6) % 7 + 3;
    let (y, _, _) = civil(thursday);
    (thursday - serial_of(y, 1, 1)).div_euclid(7) + 1
}

/// `WEEKNUM(serial, [type])`: its week of its year, the week holding 1
/// January the first, a week beginning by type 1 (the default) on a
/// Sunday, 2 on a Monday, 11 to 17 on the day `WEEKDAY`'s same type names;
/// type 21 its ISO week. A type past 21, which Excel's documentation does
/// not name, is labelled: Excel's saves answered `WEEKNUM(46023,22)` 1 and
/// then `#NUM!`, and type 22 1 for other dates in the same save. Any other
/// type `#NUM!`.
pub(crate) fn weeknum(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    answer(|| {
        if !(1..=2).contains(&args.len()) {
            return wrong();
        }
        let d = day(ctx, &args[0])?;
        let kind = match args.get(1) {
            None => 1,
            Some(a) => floored(ctx, a)?,
        };
        let start = match kind {
            1 => 0,
            2 => 1,
            11..=17 => week_start(kind),
            21 => return Ok(Value::Number(iso_week(d) as f64)),
            22.. => return label("WEEKNUM's undocumented type"),
            _ => return num(),
        };
        let (y, _, _) = civil(d);
        let first = serial_of(y, 1, 1);
        let offset = (weekday(first) - start).rem_euclid(7);
        Ok(Value::Number(
            ((d - first + offset).div_euclid(7) + 1) as f64,
        ))
    })
}

/// The functions here whose answer turns on the calendar a serial counts
/// from: every one but `DAYS`, `TIME`, `HOUR`, `MINUTE` and `SECOND`, which
/// count days and seconds alike from either.
pub const CALENDAR: [&str; 12] = [
    "DATE",
    "YEAR",
    "MONTH",
    "DAY",
    "WEEKDAY",
    "EOMONTH",
    "EDATE",
    "NETWORKDAYS",
    "WORKDAY",
    "DATEDIF",
    "YEARFRAC",
    "WEEKNUM",
];

/// The label of a function in a workbook counting its dates from 1904.
fn the_1904_system(_ctx: &Ctx<'_>, _args: &[Arg]) -> Result<Value, Reason> {
    Err(system_1904())
}

/// The reason [`label_1904`] labels by.
fn system_1904() -> Reason {
    Reason::Construct("the 1904 date system".to_string())
}

/// `TEXT` in a workbook counting its dates from 1904: its own answer, but a
/// code that draws a date labelled, since `numfmt` draws by the 1900
/// calendar.
fn text_in_1904(ctx: &Ctx<'_>, args: &[Arg]) -> Result<Value, Reason> {
    if let [_, code] = args {
        if let Ok(code) = ctx.one(code)?.text() {
            if Code::parse(&code).is_ok_and(|c| c.draws_date()) {
                return Err(system_1904());
            }
        }
    }
    super::texts::text_(ctx, args)
}

/// A library made for a workbook counting its dates from 1904, a Mac's
/// (`date1904`), which the reader reads: every function of [`CALENDAR`]
/// labelled, `the 1904 date system`, and `TEXT` of a date's code with
/// them, until a workbook asks for the system computed (KERNEL.8's
/// decision 6). A door calls it; an engine's grid counts from 1900 alone.
pub fn label_1904(library: &mut Library) {
    for name in CALENDAR {
        library.register(name, the_1904_system);
    }
    library.register("TEXT", text_in_1904);
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::calc::{library, Calc};
    use crate::sheet::{parse_a1_range, Cell, Content, Workbook};

    /// The dates the functions read, on a sheet `T`: serials of 2024's leap
    /// day, of 2026's first and of 1900's edges, a time, holidays (one on a
    /// weekend, one given twice) in B, a truth value, texts a locale reads,
    /// an error. A text is written after a `'`.
    const TABLE: &[(&str, &str)] = &[
        ("A1", "45351"),
        ("A2", "46023"),
        ("A3", "60"),
        ("A4", "0"),
        ("A5", "46023.75"),
        ("A6", "TRUE"),
        ("A7", "'1/1/2026"),
        ("A8", "'abc"),
        ("A9", "#N/A"),
        ("A10", "'46023"),
        ("B1", "46027"),
        ("B2", "46028"),
        ("B3", "46027"),
        ("B4", "46032"),
    ];

    /// Each formula's value over [`TABLE`] as `Calc` computes it, spelled.
    fn computed(formulas: &[&str]) -> Vec<String> {
        let mut wb = Workbook::new();
        let t = wb.ensure_sheet("T");
        for (addr, text) in TABLE {
            let r = parse_a1_range(addr).unwrap();
            let content = if let Some(s) = text.strip_prefix('\'') {
                Content::Text(s.to_string())
            } else if let Ok(n) = text.parse::<f64>() {
                Content::Number(n)
            } else if *text == "TRUE" || *text == "FALSE" {
                Content::Bool(*text == "TRUE")
            } else if let Some(e) = ErrorKind::parse(text) {
                Content::Error(e.text().to_string())
            } else {
                Content::Text(text.to_string())
            };
            wb.sheets[t].set(r.top, r.left, Cell { content, style: 0 });
        }
        let f = wb.ensure_sheet("F");
        for (i, formula) in formulas.iter().enumerate() {
            let content = Content::Formula(formula.to_string());
            wb.sheets[f].set(i as u32 + 1, 1, Cell { content, style: 0 });
        }
        let lib = library::language();
        let calc = Calc::run(&wb, &lib);
        (0..formulas.len())
            .map(|i| match calc.computed((f, i as u32 + 1, 1)) {
                Some(c) => c.spell(),
                None => "nothing".to_string(),
            })
            .collect()
    }

    /// Each formula of the pairs against the value it is to compute.
    fn holds(cases: &[(&str, &str)]) {
        let formulas: Vec<&str> = cases.iter().map(|(f, _)| *f).collect();
        for ((formula, want), got) in cases.iter().zip(computed(&formulas)) {
            assert_eq!(got, *want, "={formula}");
        }
    }

    const NA: &str = "(error \"#N/A\")";
    const VALUE: &str = "(error \"#VALUE!\")";
    const NUM: &str = "(error \"#NUM!\")";
    const LOCALE: &str = "(not-computed \"a locale's number\")";
    const OUTSIDE: &str = "(not-computed \"a year outside the calendar\")";
    const UNDOCUMENTED: &str = "(not-computed \"WEEKNUM's undocumented type\")";

    #[test]
    fn the_calendar_is_excel_s_1900_from_its_day_0_to_9999() {
        assert_eq!(civil(0), (1900, 1, 0));
        assert_eq!(civil(1), (1900, 1, 1));
        assert_eq!(civil(59), (1900, 2, 28));
        assert_eq!(civil(60), (1900, 2, 29));
        assert_eq!(civil(61), (1900, 3, 1));
        assert_eq!(civil(25_569), (1970, 1, 1));
        assert_eq!(civil(45_351), (2024, 2, 29));
        assert_eq!(civil(LAST_DAY), (9999, 12, 31));
        assert_eq!(civil(-1), (1899, 12, 30));
        assert_eq!(serial_of(1900, 1, 0), 0);
        assert_eq!(serial_of(1900, 2, 29), 60);
        assert_eq!(serial_of(1900, 2, 30), 61);
        assert_eq!(serial_of(1900, 3, 1), 61);
        assert_eq!(serial_of(2008, 14, 2), serial_of(2009, 2, 2));
        assert_eq!(serial_of(2008, -3, 2), serial_of(2007, 9, 2));
        assert_eq!(serial_of(2008, 1, 35), serial_of(2008, 2, 4));
        assert_eq!(serial_of(2008, 1, -15), serial_of(2007, 12, 16));
        assert_eq!(serial_of(9999, 12, 31), LAST_DAY);
        let february = |y| first_of_month(y, 3) - first_of_month(y, 2);
        assert_eq!(february(1900), 29, "February 1900, the serials'");
        assert_eq!(days_in(1900, 2), 28, "February 1900, EOMONTH's");
        assert_eq!((february(2100), days_in(2100, 2)), (28, 28));
        assert_eq!((february(2000), days_in(2000, 2)), (29, 29));
        assert_eq!(
            (days_in(2026, 1), days_in(2026, 4), days_in(2026, 12)),
            (31, 30, 31)
        );
        // Every day from 0 to the last: one after another, each the serial
        // of its own date, a Sunday every seventh from serial 1.
        let mut was = civil(0);
        for s in 1..=LAST_DAY {
            let now = civil(s);
            assert!(was < now, "{s}");
            assert_eq!(serial_of(now.0, now.1, now.2), s, "{s}");
            was = now;
        }
        assert_eq!(weekday(1), 0);
        assert_eq!(weekday(0), 6);
    }

    #[test]
    fn a_date_is_made_and_read_apart() {
        holds(&[
            ("DATE(2026,1,1)", "46023"),
            ("DATE(1900,1,0)", "0"),
            ("DATE(1900,2,29)", "60"),
            ("DATE(108,1,2)", "39449"),
            ("DATE(1899,12,31)", "693962"),
            ("DATE(0,1,1)", "1"),
            ("DATE(2008,14,2)", "39846"),
            ("DATE(2008,-3,2)", "39327"),
            ("DATE(2008,1,-15)", "39432"),
            ("DATE(1900,1,-1)", NUM),
            ("DATE(1900,0,31)", NUM),
            ("DATE(-1,1,1)", NUM),
            ("DATE(10000,1,1)", NUM),
            ("DATE(10000,-11,1)", OUTSIDE),
            ("DATE(-1,13,1)", OUTSIDE),
            ("DATE(-1,12,1)", NUM),
            ("DATE(10000,-12,1)", OUTSIDE),
            ("DATE(10000,1,-400)", NUM),
            ("DATE(1899,13,1)", "693963"),
            ("DATE(9999,13,-31)", NUM),
            ("DATE(9999,12,31)", "2958465"),
            ("DATE(9999,12,32)", NUM),
            ("DATE(2026.9,1.9,1.9)", "46023"),
            ("DATE(2026,1,-0.5)", "46021"),
            ("DATE(2026,-0.5,1)", "45962"),
            ("DATE(-0.5,1,1)", NUM),
            ("DATE(T!A9,1,1)", NA),
            ("YEAR(T!A1)", "2024"),
            ("MONTH(T!A1)", "2"),
            ("DAY(T!A1)", "29"),
            ("YEAR(T!A3)", "1900"),
            ("DAY(T!A3)", "29"),
            ("DAY(T!A4)", "0"),
            ("MONTH(T!A4)", "1"),
            ("DAY(T!A5)", "1"),
            ("YEAR(T!A6)", "1900"),
            ("YEAR(T!A7)", LOCALE),
            ("YEAR(T!A8)", VALUE),
            ("YEAR(T!A10)", "2026"),
            ("YEAR(-1)", NUM),
            ("YEAR(2958466)", NUM),
            ("YEAR(2958465.9)", "9999"),
            ("YEAR(2958465.9999999)", NUM),
            ("YEAR(-0.000001)", NUM),
            ("WEEKDAY(45658.99999999)", "5"),
            ("DAY(45658.99999)", "1"),
            ("DAY(45658.99999999)", "2"),
            ("WEEKDAY(T!A2)", "5"),
            ("WEEKDAY(T!A2,2)", "4"),
            ("WEEKDAY(T!A2,3)", "3"),
            ("WEEKDAY(T!A2,11)", "4"),
            ("WEEKDAY(T!A2,14)", "1"),
            ("WEEKDAY(T!A2,17)", "5"),
            ("WEEKDAY(1)", "1"),
            ("WEEKDAY(0)", "7"),
            ("WEEKDAY(T!A2,4)", NUM),
        ]);
    }

    #[test]
    fn a_month_is_moved_and_its_end_found() {
        holds(&[
            ("EOMONTH(T!A2,0)", "46053"),
            ("EOMONTH(T!A2,1)", "46081"),
            ("EOMONTH(T!A1,-1)", "45322"),
            ("EOMONTH(T!A1,12)", "45716"),
            ("EOMONTH(T!A2,-1.9)", "46022"),
            ("EOMONTH(DATE(1900,1,15),1)", "59"),
            ("EOMONTH(60,0)", "59"),
            ("EOMONTH(0,0)", "31"),
            ("EOMONTH(1,-1)", NUM),
            ("EOMONTH(DATE(9999,12,1),1)", NUM),
            ("EDATE(DATE(2011,1,31),1)", "40602"),
            ("EDATE(T!A1,12)", "45716"),
            ("EDATE(T!A1,48)", "46812"),
            ("EDATE(T!A2,-1.9)", "45992"),
            ("EDATE(DATE(1900,1,31),1)", "59"),
            ("EDATE(60,12)", "425"),
            ("EDATE(0,0)", "0"),
            ("EDATE(0,1)", "31"),
            ("EDATE(31,-1)", NUM),
            ("EDATE(60,0)", "59"),
            ("EOMONTH(45688.99999999,0)", "45688"),
            ("EDATE(45658.99999999,0)", "45658"),
            ("DAYS(45658.99999999,45658)", "0"),
            ("DAYS(T!A2,T!A1)", "672"),
            ("DAYS(T!A1,T!A2)", "-672"),
            ("DAYS(T!A5,T!A2)", "0"),
            ("DAYS(T!A7,T!A2)", LOCALE),
        ]);
    }

    #[test]
    fn working_days_pass_weekends_and_holidays_by() {
        // 2026-01-01 is a Thursday; its holidays, a Monday given twice, a
        // Tuesday and a Saturday.
        holds(&[
            ("NETWORKDAYS(T!A2,T!A2+13)", "10"),
            ("NETWORKDAYS(T!A2,T!A2+13,T!B1:B4)", "8"),
            ("NETWORKDAYS(T!A2+13,T!A2,T!B1:B4)", "-8"),
            ("NETWORKDAYS(T!A2+2,T!A2+3)", "0"),
            ("NETWORKDAYS(T!A2,T!A2,T!A2)", "0"),
            ("NETWORKDAYS(T!A2,T!A2+1,T!A9)", NA),
            ("NETWORKDAYS(T!A2,T!A2+1,TRUE)", VALUE),
            ("WORKDAY(T!A2,1)", "46024"),
            ("WORKDAY(T!A2,2)", "46027"),
            ("WORKDAY(T!A2,2,T!B1:B4)", "46029"),
            ("WORKDAY(T!A2,-1)", "46022"),
            ("WORKDAY(T!A2,1.9)", "46024"),
            ("WORKDAY(T!A2,-1.9)", "46021"),
            ("WORKDAY(T!A2,0)", "46023"),
            ("WORKDAY(T!A2+2,0)", "46025"),
            ("WORKDAY(1,-1)", NUM),
        ]);
    }

    #[test]
    fn a_time_is_a_day_s_fraction_read_to_the_second() {
        holds(&[
            ("TIME(18,0,0)", "0.75"),
            ("TIME(24,0,0)", "0"),
            ("TIME(1,-30,0)", "0.020833333333333332"),
            ("TIME(25,0,0)", "0.04166666666666674"),
            ("TIME(32767,0,0)", "0.29166666666674246"),
            ("TIME(0,0,59.9)", "0.0006828703703703704"),
            ("TIME(0,0,-1)", NUM),
            ("TIME(0,0,-0.5)", NUM),
            ("TIME(1,-0.5,0)", "0.04097222222222222"),
            ("TIME(-0.5,60,0)", "0"),
            ("TIME(32768,0,0)", NUM),
            ("HOUR(T!A5)", "18"),
            ("MINUTE(T!A5)", "0"),
            ("HOUR(0.999999)", "0"),
            ("HOUR(0.99999)", "23"),
            ("SECOND(2958465.99999)", "59"),
            ("SECOND(2958465.9999999)", NUM),
            ("SECOND(TIME(1,2,3))", "3"),
            ("MINUTE(TIME(1,2,3))", "2"),
            ("SECOND(T!A7)", LOCALE),
        ]);
    }

    #[test]
    fn spans_of_dates_are_counted_by_their_units_and_bases() {
        holds(&[
            ("DATEDIF(DATE(2001,1,1),DATE(2003,1,1),\"Y\")", "2"),
            ("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"d\")", "440"),
            ("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"M\")", "14"),
            ("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YM\")", "2"),
            ("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"YD\")", "75"),
            ("DATEDIF(DATE(2001,6,1),DATE(2002,8,15),\"MD\")", "14"),
            ("DATEDIF(DATE(2011,1,31),DATE(2011,3,1),\"MD\")", "-2"),
            ("DATEDIF(DATE(2019,3,1),DATE(2020,2,29),\"YD\")", "365"),
            ("DATEDIF(DATE(2020,2,29),DATE(2021,3,1),\"YD\")", "0"),
            ("DATEDIF(DATE(2020,2,29),DATE(2021,2,28),\"YD\")", "365"),
            ("DATEDIF(DATE(2020,2,29),DATE(2022,2,28),\"YD\")", "365"),
            ("DATEDIF(DATE(2019,2,28),DATE(2020,2,29),\"YD\")", "1"),
            ("DATEDIF(DATE(1988,6,22),DATE(2012,5,11),\"YD\")", "323"),
            ("DATEDIF(DATE(2018,12,31),DATE(2019,1,1),\"YD\")", "1"),
            ("DATEDIF(DATE(2001,6,1),DATE(2003,5,1),\"YD\")", "334"),
            ("DATEDIF(DATE(2019,3,1),DATE(2021,2,28),\"YD\")", "364"),
            ("DATEDIF(DATE(2016,2,29),DATE(2018,1,15),\"YD\")", "321"),
            ("DATEDIF(DATE(2020,2,29),DATE(2023,3,1),\"YD\")", "0"),
            ("DATEDIF(DATE(2011,1,31),DATE(2011,3,1),\"YD\")", "29"),
            ("DATEDIF(DATE(2020,2,29),DATE(2021,2,28),\"MD\")", "30"),
            ("DATEDIF(DATE(2019,12,31),DATE(2020,1,1),\"YD\")", "1"),
            ("DATEDIF(DATE(2001,6,1),DATE(2002,6,1),\"YD\")", "0"),
            ("DATEDIF(DATE(2003,1,1),DATE(2001,1,1),\"Y\")", NUM),
            ("DATEDIF(DATE(2001,1,1),DATE(2003,1,1),\"W\")", NUM),
            (
                "YEARFRAC(DATE(2012,1,1),DATE(2012,7,30))",
                "0.5805555555555556",
            ),
            (
                "YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),1)",
                "0.5765027322404371",
            ),
            (
                "YEARFRAC(DATE(2012,1,1),DATE(2012,7,30),3)",
                "0.5780821917808219",
            ),
            (
                "YEARFRAC(DATE(2012,7,30),DATE(2012,1,1),2)",
                "0.5861111111111111",
            ),
            (
                "YEARFRAC(DATE(2012,1,31),DATE(2012,3,31),4)",
                "0.16666666666666666",
            ),
            ("YEARFRAC(DATE(1900,1,1),DATE(1900,12,31),1)", "1"),
            (
                "YEARFRAC(DATE(1900,1,1),DATE(1901,1,1),1)",
                "1.0027397260273974",
            ),
            (
                "YEARFRAC(DATE(1900,1,1),DATE(1902,1,2),1)",
                "2.0054794520547947",
            ),
            (
                "YEARFRAC(DATE(1900,2,28),DATE(1900,3,31),0)",
                "0.08611111111111111",
            ),
            ("YEARFRAC(60,DATE(1900,3,31),0)", "0.08888888888888889"),
            ("YEARFRAC(T!A1,T!A2,-0.5)", NUM),
            (
                "YEARFRAC(DATE(2010,6,15),DATE(2013,2,1),1)",
                "2.6338124572210813",
            ),
            (
                "YEARFRAC(DATE(2012,2,29),DATE(2013,1,1),1)",
                "0.8387978142076503",
            ),
            (
                "YEARFRAC(DATE(2011,1,30),DATE(2011,3,31),0)",
                "0.16666666666666666",
            ),
            ("YEARFRAC(T!A2,T!A2)", "0"),
            ("YEARFRAC(T!A1,T!A2,5)", NUM),
            ("WEEKNUM(T!A2)", "1"),
            ("WEEKNUM(DATE(2026,1,4))", "2"),
            ("WEEKNUM(DATE(2026,1,4),2)", "1"),
            ("WEEKNUM(DATE(2026,1,5),2)", "2"),
            ("WEEKNUM(DATE(2026,1,1),21)", "1"),
            ("WEEKNUM(DATE(2027,1,1),21)", "53"),
            ("WEEKNUM(DATE(2024,12,30),21)", "1"),
            ("WEEKNUM(1,21)", "52"),
            ("WEEKNUM(0)", "0"),
            ("WEEKNUM(T!A2,22)", UNDOCUMENTED),
            ("WEEKNUM(T!A2,150)", UNDOCUMENTED),
            ("WEEKNUM(T!A2,20)", NUM),
            ("WEEKNUM(T!A2,0)", NUM),
            ("WEEKNUM(45661.99999999)", "1"),
            ("WEEKNUM(T!A2,3)", NUM),
        ]);
    }

    #[test]
    fn a_workbook_of_1904_labels_the_calendar_s_functions() {
        let mut lib = library::language();
        label_1904(&mut lib);
        let mut wb = Workbook::new();
        let f = wb.ensure_sheet("F");
        let formulas = [
            ("YEAR(1)", "(not-computed \"the 1904 date system\")"),
            ("DAYS(10,1)", "9"),
            ("HOUR(0.5)", "12"),
            (
                "TEXT(1,\"yyyy\")",
                "(not-computed \"the 1904 date system\")",
            ),
            ("TEXT(1.5,\"0.0\")", "\"1.5\""),
        ];
        for (i, (formula, _)) in formulas.iter().enumerate() {
            let content = Content::Formula(formula.to_string());
            wb.sheets[f].set(i as u32 + 1, 1, Cell { content, style: 0 });
        }
        let calc = Calc::run(&wb, &lib);
        for (i, (formula, want)) in formulas.iter().enumerate() {
            let got = calc.computed((f, i as u32 + 1, 1)).unwrap().spell();
            assert_eq!(got, *want, "={formula}");
        }
    }
}
