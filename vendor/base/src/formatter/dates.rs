use chrono::Datelike;
use chrono::Duration;
use chrono::Months;
use chrono::NaiveDate;
use chrono::Weekday;

use crate::constants::EXCEL_DATE_BASE;
use crate::constants::MAXIMUM_DATE_SERIAL_NUMBER;
use crate::constants::MINIMUM_DATE_SERIAL_NUMBER;

pub const DATE_OUT_OF_RANGE_MESSAGE: &str = "Out of range parameters for date";

// Excel inherited Lotus 1-2-3's belief that 1900 was a leap year: serial 60
// is the fictitious 1900-02-29. Real dates before 1900-03-01 are therefore
// one serial lower than the naive EXCEL_DATE_BASE offset suggests
// (serial 1 = 1900-01-01, serial 59 = 1900-02-28, serial 61 = 1900-03-01).

#[inline]
pub(crate) fn convert_to_serial_number(date: NaiveDate) -> i32 {
    let serial = date.num_days_from_ce() - EXCEL_DATE_BASE;
    if serial < 61 {
        serial - 1
    } else {
        serial
    }
}

fn is_date_within_range(date: NaiveDate) -> bool {
    convert_to_serial_number(date) >= MINIMUM_DATE_SERIAL_NUMBER
        && convert_to_serial_number(date) <= MAXIMUM_DATE_SERIAL_NUMBER
}

/// The real calendar date for an Excel serial number. The fictitious serial
/// 60 maps to 1900-02-28; use [`excel_date_components`] where the year/month/
/// day Excel reports (1900-02-29) matters.
pub fn from_excel_date(days: i64) -> Result<NaiveDate, String> {
    if days < MINIMUM_DATE_SERIAL_NUMBER as i64 {
        return Err(format!(
            "Excel date must be greater than {MINIMUM_DATE_SERIAL_NUMBER}"
        ));
    };
    if days > MAXIMUM_DATE_SERIAL_NUMBER as i64 {
        return Err(format!(
            "Excel date must be less than {MAXIMUM_DATE_SERIAL_NUMBER}"
        ));
    };
    #[allow(clippy::expect_used)]
    let dt = NaiveDate::from_ymd_opt(1900, 1, 1).expect("problem with chrono::NaiveDate");
    if days <= 59 {
        Ok(dt + Duration::days(days - 1))
    } else {
        Ok(dt + Duration::days(days - 2))
    }
}

/// The (year, month, day) Excel reports for a serial number, honoring
/// serial 0 (the phantom 1900-01-00) and serial 60 (1900-02-29).
pub fn excel_date_components(serial: i64) -> Result<(i32, u32, u32), String> {
    if serial == 0 {
        return Ok((1900, 1, 0));
    }
    if serial == 60 {
        return Ok((1900, 2, 29));
    }
    let date = from_excel_date(serial)?;
    Ok((date.year(), date.month(), date.day()))
}

/// Excel's weekday cycle is continuous over serial numbers — the fictitious
/// 1900-02-29 consumes a slot — so serial 1 is a Sunday even though
/// 1900-01-01 was really a Monday. From serial 61 (1900-03-01) onward this
/// agrees with the calendar weekday.
pub fn excel_weekday_from_serial(serial: i64) -> Weekday {
    match serial.rem_euclid(7) {
        0 => Weekday::Sat,
        1 => Weekday::Sun,
        2 => Weekday::Mon,
        3 => Weekday::Tue,
        4 => Weekday::Wed,
        5 => Weekday::Thu,
        _ => Weekday::Fri,
    }
}

pub fn date_to_serial_number(day: u32, month: u32, year: i32) -> Result<i32, String> {
    match NaiveDate::from_ymd_opt(year, month, day) {
        Some(native_date) => Ok(convert_to_serial_number(native_date)),
        None => Err(DATE_OUT_OF_RANGE_MESSAGE.to_string()),
    }
}

pub fn permissive_date_to_serial_number(day: i32, month: i32, year: i32) -> Result<i32, String> {
    // Excel parses `DATE` very permissively. It allows not just for valid date values, but it
    // allows for invalid dates as well. If you for example enter `DATE(1900, 1, 32)` it will
    // return the date `1900-02-01`. Despite giving a day that is out of range it will just
    // wrap the month and year around.
    //
    // This function applies that same logic to dates. And does it in the most compatible way as
    // possible.

    let Some(mut date) = NaiveDate::from_ymd_opt(year, 1, 1) else {
        return Err(DATE_OUT_OF_RANGE_MESSAGE.to_string());
    };

    // One thing to note for example is that even if you started with a year out of range
    // but tried to increment the months so that it wraps around into within range, excel
    // would still return an error.
    //
    // I.E. DATE(0,13,-1) will return an error, despite it being equivalent to DATE(1,1,0) which
    // is within range.
    //
    // As a result, we have to run range checks as we parse the date from the biggest unit to the
    // smallest unit.
    if !is_date_within_range(date) {
        return Err(DATE_OUT_OF_RANGE_MESSAGE.to_string());
    }

    date = {
        let month_diff = month - 1;
        let abs_month = month_diff.unsigned_abs();
        if month_diff <= 0 {
            date = date - Months::new(abs_month);
        } else {
            date = date + Months::new(abs_month);
        }
        if !is_date_within_range(date) {
            return Err(DATE_OUT_OF_RANGE_MESSAGE.to_string());
        }
        date
    };

    // Day offsets are serial-number arithmetic in Excel's fictional calendar:
    // DATE(1900,2,29) is the phantom serial 60, and DATE(1900,1,60) too.
    let serial = convert_to_serial_number(date) as i64 + (day as i64 - 1);
    if !(MINIMUM_DATE_SERIAL_NUMBER as i64..=MAXIMUM_DATE_SERIAL_NUMBER as i64).contains(&serial) {
        return Err(DATE_OUT_OF_RANGE_MESSAGE.to_string());
    }
    Ok(serial as i32)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_permissive_date_to_serial_number() {
        assert_eq!(
            permissive_date_to_serial_number(42, 42, 2002),
            date_to_serial_number(12, 7, 2005)
        );
        assert_eq!(
            permissive_date_to_serial_number(1, 42, 2002),
            date_to_serial_number(1, 6, 2005)
        );
        assert_eq!(
            permissive_date_to_serial_number(1, 15, 2000),
            date_to_serial_number(1, 3, 2001)
        );
        assert_eq!(
            permissive_date_to_serial_number(1, 49, 2000),
            date_to_serial_number(1, 1, 2004)
        );
        assert_eq!(
            permissive_date_to_serial_number(1, 49, 2000),
            date_to_serial_number(1, 1, 2004)
        );
        assert_eq!(
            permissive_date_to_serial_number(31, 49, 2000),
            date_to_serial_number(31, 1, 2004)
        );
        assert_eq!(
            permissive_date_to_serial_number(256, 49, 2000),
            date_to_serial_number(12, 9, 2004)
        );
        assert_eq!(
            permissive_date_to_serial_number(256, 1, 2004),
            date_to_serial_number(12, 9, 2004)
        );
    }

    #[test]
    fn test_max_and_min_dates() {
        assert_eq!(
            permissive_date_to_serial_number(31, 12, 9999),
            Ok(MAXIMUM_DATE_SERIAL_NUMBER),
        );
        assert_eq!(
            permissive_date_to_serial_number(1, 1, 1900),
            Ok(MINIMUM_DATE_SERIAL_NUMBER),
        );
        // Before the epoch: no serial number
        assert!(permissive_date_to_serial_number(31, 12, 1899).is_err());
    }

    #[test]
    fn test_1900_leap_bug_serials() {
        use chrono::NaiveDate;
        // Excel: serial 1 = 1900-01-01, 59 = 1900-02-28, 61 = 1900-03-01
        assert_eq!(date_to_serial_number(1, 1, 1900), Ok(1));
        assert_eq!(date_to_serial_number(28, 2, 1900), Ok(59));
        assert_eq!(date_to_serial_number(1, 3, 1900), Ok(61));
        assert_eq!(
            from_excel_date(1),
            Ok(NaiveDate::from_ymd_opt(1900, 1, 1).unwrap())
        );
        assert_eq!(
            from_excel_date(59),
            Ok(NaiveDate::from_ymd_opt(1900, 2, 28).unwrap())
        );
        assert_eq!(
            from_excel_date(61),
            Ok(NaiveDate::from_ymd_opt(1900, 3, 1).unwrap())
        );
        // The phantom day
        assert_eq!(excel_date_components(60), Ok((1900, 2, 29)));
        assert_eq!(excel_date_components(0), Ok((1900, 1, 0)));
        // Modern dates keep their serials
        assert_eq!(date_to_serial_number(1, 1, 2020), Ok(43831));
        assert_eq!(
            from_excel_date(43831),
            Ok(NaiveDate::from_ymd_opt(2020, 1, 1).unwrap())
        );
        // Excel says serial 1 is a Sunday, serial 61 (1900-03-01) a Thursday
        assert_eq!(excel_weekday_from_serial(1), Weekday::Sun);
        assert_eq!(excel_weekday_from_serial(61), Weekday::Thu);
        assert_eq!(excel_weekday_from_serial(60), Weekday::Wed);
    }
}
