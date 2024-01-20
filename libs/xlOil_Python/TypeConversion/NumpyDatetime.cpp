#include "NumpyDatetime.h"
#include <stdexcept>
/*
 * This file implements core functionality for NumPy datetime.
 *
 * Written by Mark Wiebe (mwwiebe@gmail.com)
 * Copyright (c) 2011 by Enthought, Inc.
 *
 * See LICENSE.txt for the license.
 */

/*
 * This function returns a pointer to the DateTimeMetaData
 * contained within the provided datetime dtype.
 */
PyArray_DatetimeMetaData *
get_datetime_metadata_from_dtype(PyArray_Descr *dtype)
{
    if (!PyDataType_ISDATETIME(dtype))
      throw std::runtime_error("cannot get datetime metadata from non-datetime type");

    return &(((PyArray_DatetimeDTypeMetaData *)dtype->c_metadata)->meta);
}

PyArray_Descr *
create_datetime_dtype(int type_num, PyArray_DatetimeMetaData *meta)
{
    PyArray_Descr *dtype = NULL;
    PyArray_DatetimeMetaData *dt_data;

    /* Create a default datetime or timedelta */
    if (type_num == NPY_DATETIME || type_num == NPY_TIMEDELTA) {
        dtype = PyArray_DescrNewFromType(type_num);
    }
    else {
        throw std::runtime_error("Asked to create a datetime type with a non-datetime type number");
    }

    if (dtype == NULL) {
        return NULL;
    }

    dt_data = &(((PyArray_DatetimeDTypeMetaData *)dtype->c_metadata)->meta);

    /* Copy the metadata */
    *dt_data = *meta;

    return dtype;
}

/*
 * Computes the python `ret, d = divmod(d, unit)`.
 *
 * Note that GCC is smart enough at -O2 to eliminate the `if(*d < 0)` branch
 * for subsequent calls to this command - it is able to deduce that `*d >= 0`.
 */
inline
npy_int64 extract_unit_64(npy_int64 *d, npy_int64 unit) {
    assert(unit > 0);
    npy_int64 div = *d / unit;
    npy_int64 mod = *d % unit;
    if (mod < 0) {
        mod += unit;
        div -= 1;
    }
    assert(mod >= 0);
    *d = mod;
    return div;
}

inline
npy_int32 extract_unit_32(npy_int32 *d, npy_int32 unit) {
    assert(unit > 0);
    npy_int32 div = *d / unit;
    npy_int32 mod = *d % unit;
    if (mod < 0) {
        mod += unit;
        div -= 1;
    }
    assert(mod >= 0);
    *d = mod;
    return div;
}

/* Days per month, regular year and leap year */
int _days_per_month_table[2][12] = {
    { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 },
    { 31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 }
};

/*
 * Returns 1 if the given year is a leap year, 0 otherwise.
 */
int
is_leapyear(npy_int64 year)
{
    return (year & 0x3) == 0 && /* year % 4 == 0 */
           ((year % 100) != 0 ||
            (year % 400) == 0);
}

/*
 * Calculates the days offset from the 1970 epoch.
 */
npy_int64
get_datetimestruct_days(const npy_datetimestruct *dts)
{
    int i, month;
    npy_int64 year, days = 0;
    int *month_lengths;

    year = dts->year - 1970;
    days = year * 365;

    /* Adjust for leap years */
    if (days >= 0) {
        /*
         * 1968 is the closest leap year before 1970.
         * Exclude the current year, so add 1.
         */
        year += 1;
        /* Add one day for each 4 years */
        days += year / 4;
        /* 1900 is the closest previous year divisible by 100 */
        year += 68;
        /* Subtract one day for each 100 years */
        days -= year / 100;
        /* 1600 is the closest previous year divisible by 400 */
        year += 300;
        /* Add one day for each 400 years */
        days += year / 400;
    }
    else {
        /*
         * 1972 is the closest later year after 1970.
         * Include the current year, so subtract 2.
         */
        year -= 2;
        /* Subtract one day for each 4 years */
        days += year / 4;
        /* 2000 is the closest later year divisible by 100 */
        year -= 28;
        /* Add one day for each 100 years */
        days -= year / 100;
        /* 2000 is also the closest later year divisible by 400 */
        /* Subtract one day for each 400 years */
        days += year / 400;
    }

    month_lengths = _days_per_month_table[is_leapyear(dts->year)];
    month = dts->month - 1;

    /* Add the months */
    for (i = 0; i < month; ++i) {
        days += month_lengths[i];
    }

    /* Add the days */
    days += dts->day - 1;

    return days;
}

/*
 * Calculates the minutes offset from the 1970 epoch.
 */
npy_int64
get_datetimestruct_minutes(const npy_datetimestruct *dts)
{
    npy_int64 days = get_datetimestruct_days(dts) * 24 * 60;
    days += dts->hour * 60;
    days += dts->min;

    return days;
}

/*
 * Modifies '*days_' to be the day offset within the year,
 * and returns the year.
 */
npy_int64
days_to_yearsdays(npy_int64 *days_)
{
    const npy_int64 days_per_400years = (400*365 + 100 - 4 + 1);
    /* Adjust so it's relative to the year 2000 (divisible by 400) */
    npy_int64 days = (*days_) - (365*30 + 7);
    npy_int64 year;

    /* Break down the 400 year cycle to get the year and day within the year */
    year = 400 * extract_unit_64(&days, days_per_400years);

    /* Work out the year/day within the 400 year cycle */
    if (days >= 366) {
        year += 100 * ((days-1) / (100*365 + 25 - 1));
        days = (days-1) % (100*365 + 25 - 1);
        if (days >= 365) {
            year += 4 * ((days+1) / (4*365 + 1));
            days = (days+1) % (4*365 + 1);
            if (days >= 366) {
                year += (days-1) / 365;
                days = (days-1) % 365;
            }
        }
    }

    *days_ = days;
    return year + 2000;
}

/* Extracts the month number from a 'datetime64[D]' value */
int
days_to_month_number(npy_datetime days)
{
    npy_int64 year;
    int *month_lengths, i;

    year = days_to_yearsdays(&days);
    month_lengths = _days_per_month_table[is_leapyear(year)];

    for (i = 0; i < 12; ++i) {
        if (days < month_lengths[i]) {
            return i + 1;
        }
        else {
            days -= month_lengths[i];
        }
    }

    /* Should never get here */
    return 1;
}

/*
 * Fills in the year, month, day in 'dts' based on the days
 * offset from 1970.
 */
void
set_datetimestruct_days(npy_int64 days, npy_datetimestruct *dts)
{
    int *month_lengths, i;

    dts->year = days_to_yearsdays(&days);
    month_lengths = _days_per_month_table[is_leapyear(dts->year)];

    for (i = 0; i < 12; ++i) {
        if (days < month_lengths[i]) {
            dts->month = i + 1;
            dts->day = (int)days + 1;
            return;
        }
        else {
            days -= month_lengths[i];
        }
    }
}

/*NUMPY_API
 *
 * Converts a datetime from a datetimestruct to a datetime based
 * on some metadata. The date is assumed to be valid.
 *
 * TODO: If meta->num is really big, there could be overflow
 *
 * Returns 0 on success, -1 on failure.
 */
int
NpyDatetime_ConvertDatetimeStructToDatetime64(PyArray_DatetimeMetaData *meta,
                                    const npy_datetimestruct *dts,
                                    npy_datetime *out)
{
    npy_datetime ret;
    NPY_DATETIMEUNIT base = meta->base;

    /* If the datetimestruct is NaT, return NaT */
    if (dts->year == NPY_DATETIME_NAT) {
        *out = NPY_DATETIME_NAT;
        return 0;
    }

    /* Cannot instantiate a datetime with generic units */
    if (meta->base == NPY_FR_GENERIC) {
        throw std::runtime_error(
                    "Cannot create a NumPy datetime other than NaT with generic units");
    }

    if (base == NPY_FR_Y) {
        /* Truncate to the year */
        ret = dts->year - 1970;
    }
    else if (base == NPY_FR_M) {
        /* Truncate to the month */
        ret = 12 * (dts->year - 1970) + (dts->month - 1);
    }
    else {
        /* Otherwise calculate the number of days to start */
        npy_int64 days = get_datetimestruct_days(dts);

        switch (base) {
            case NPY_FR_W:
                /* Truncate to weeks */
                if (days >= 0) {
                    ret = days / 7;
                }
                else {
                    ret = (days - 6) / 7;
                }
                break;
            case NPY_FR_D:
                ret = days;
                break;
            case NPY_FR_h:
                ret = days * 24 +
                      dts->hour;
                break;
            case NPY_FR_m:
                ret = (days * 24 +
                      dts->hour) * 60 +
                      dts->min;
                break;
            case NPY_FR_s:
                ret = ((days * 24 +
                      dts->hour) * 60 +
                      dts->min) * 60 +
                      dts->sec;
                break;
            case NPY_FR_ms:
                ret = (((days * 24 +
                      dts->hour) * 60 +
                      dts->min) * 60 +
                      dts->sec) * 1000 +
                      dts->us / 1000;
                break;
            case NPY_FR_us:
                ret = (((days * 24 +
                      dts->hour) * 60 +
                      dts->min) * 60 +
                      dts->sec) * 1000000 +
                      dts->us;
                break;
            case NPY_FR_ns:
                ret = ((((days * 24 +
                      dts->hour) * 60 +
                      dts->min) * 60 +
                      dts->sec) * 1000000 +
                      dts->us) * 1000 +
                      dts->ps / 1000;
                break;
            case NPY_FR_ps:
                ret = ((((days * 24 +
                      dts->hour) * 60 +
                      dts->min) * 60 +
                      dts->sec) * 1000000 +
                      dts->us) * 1000000 +
                      dts->ps;
                break;
            case NPY_FR_fs:
                /* only 2.6 hours */
                ret = (((((days * 24 +
                      dts->hour) * 60 +
                      dts->min) * 60 +
                      dts->sec) * 1000000 +
                      dts->us) * 1000000 +
                      dts->ps) * 1000 +
                      dts->as / 1000;
                break;
            case NPY_FR_as:
                /* only 9.2 secs */
                ret = (((((days * 24 +
                      dts->hour) * 60 +
                      dts->min) * 60 +
                      dts->sec) * 1000000 +
                      dts->us) * 1000000 +
                      dts->ps) * 1000000 +
                      dts->as;
                break;
            default:
                /* Something got corrupted */
                throw std::runtime_error("NumPy datetime metadata with corrupt unit value");
        }
    }

    /* Divide by the multiplier */
    if (meta->num > 1) {
        if (ret >= 0) {
            ret /= meta->num;
        }
        else {
            ret = (ret - meta->num + 1) / meta->num;
        }
    }

    *out = ret;

    return 0;
}

/*NUMPY_API
 *
 * Converts a datetime based on the given metadata into a datetimestruct
 */
int
NpyDatetime_ConvertDatetime64ToDatetimeStruct(
        PyArray_DatetimeMetaData *meta, npy_datetime dt,
        npy_datetimestruct *out)
{
    npy_int64 days;

    /* Initialize the output to all zeros */
    memset(out, 0, sizeof(npy_datetimestruct));
    out->year = 1970;
    out->month = 1;
    out->day = 1;

    /* NaT is signaled in the year */
    if (dt == NPY_DATETIME_NAT) {
        out->year = NPY_DATETIME_NAT;
        return 0;
    }

    /* Datetimes can't be in generic units */
    if (meta->base == NPY_FR_GENERIC) {
        throw std::runtime_error(
                    "Cannot convert a NumPy datetime value other than NaT "
                    "with generic units");
    }

    /* TODO: Change to a mechanism that avoids the potential overflow */
    dt *= meta->num;

    /*
     * Note that care must be taken with the / and % operators
     * for negative values.
     */
    switch (meta->base) {
        case NPY_FR_Y:
            out->year = 1970 + dt;
            break;

        case NPY_FR_M:
            out->year  = 1970 + extract_unit_64(&dt, 12);
            out->month = dt + 1;
            break;

        case NPY_FR_W:
            /* A week is 7 days */
            set_datetimestruct_days(dt * 7, out);
            break;

        case NPY_FR_D:
            set_datetimestruct_days(dt, out);
            break;

        case NPY_FR_h:
            days      = extract_unit_64(&dt, 24LL);
            set_datetimestruct_days(days, out);
            out->hour = (int)dt;
            break;

        case NPY_FR_m:
            days      =      extract_unit_64(&dt, 60LL*24);
            set_datetimestruct_days(days, out);
            out->hour = (int)extract_unit_64(&dt, 60LL);
            out->min  = (int)dt;
            break;

        case NPY_FR_s:
            days      =      extract_unit_64(&dt, 60LL*60*24);
            set_datetimestruct_days(days, out);
            out->hour = (int)extract_unit_64(&dt, 60LL*60);
            out->min  = (int)extract_unit_64(&dt, 60LL);
            out->sec  = (int)dt;
            break;

        case NPY_FR_ms:
            days      =      extract_unit_64(&dt, 1000LL*60*60*24);
            set_datetimestruct_days(days, out);
            out->hour = (int)extract_unit_64(&dt, 1000LL*60*60);
            out->min  = (int)extract_unit_64(&dt, 1000LL*60);
            out->sec  = (int)extract_unit_64(&dt, 1000LL);
            out->us   = (int)(dt * 1000);
            break;

        case NPY_FR_us:
            days      =      extract_unit_64(&dt, 1000LL*1000*60*60*24);
            set_datetimestruct_days(days, out);
            out->hour = (int)extract_unit_64(&dt, 1000LL*1000*60*60);
            out->min  = (int)extract_unit_64(&dt, 1000LL*1000*60);
            out->sec  = (int)extract_unit_64(&dt, 1000LL*1000);
            out->us   = (int)dt;
            break;

        case NPY_FR_ns:
            days      =      extract_unit_64(&dt, 1000LL*1000*1000*60*60*24);
            set_datetimestruct_days(days, out);
            out->hour = (int)extract_unit_64(&dt, 1000LL*1000*1000*60*60);
            out->min  = (int)extract_unit_64(&dt, 1000LL*1000*1000*60);
            out->sec  = (int)extract_unit_64(&dt, 1000LL*1000*1000);
            out->us   = (int)extract_unit_64(&dt, 1000LL);
            out->ps   = (int)(dt * 1000);
            break;

        case NPY_FR_ps:
            days      =      extract_unit_64(&dt, 1000LL*1000*1000*1000*60*60*24);
            set_datetimestruct_days(days, out);
            out->hour = (int)extract_unit_64(&dt, 1000LL*1000*1000*1000*60*60);
            out->min  = (int)extract_unit_64(&dt, 1000LL*1000*1000*1000*60);
            out->sec  = (int)extract_unit_64(&dt, 1000LL*1000*1000*1000);
            out->us   = (int)extract_unit_64(&dt, 1000LL*1000);
            out->ps   = (int)(dt);
            break;

        case NPY_FR_fs:
            /* entire range is only +- 2.6 hours */
            out->hour = (int)extract_unit_64(&dt, 1000LL*1000*1000*1000*1000*60*60);
            if (out->hour < 0) {
                out->year  = 1969;
                out->month = 12;
                out->day   = 31;
                out->hour  += 24;
                assert(out->hour >= 0);
            }
            out->min  = (int)extract_unit_64(&dt, 1000LL*1000*1000*1000*1000*60);
            out->sec  = (int)extract_unit_64(&dt, 1000LL*1000*1000*1000*1000);
            out->us   = (int)extract_unit_64(&dt, 1000LL*1000*1000);
            out->ps   = (int)extract_unit_64(&dt, 1000LL);
            out->as   = (int)(dt * 1000);
            break;

        case NPY_FR_as:
            /* entire range is only +- 9.2 seconds */
            out->sec = (int)extract_unit_64(&dt, 1000LL*1000*1000*1000*1000*1000);
            if (out->sec < 0) {
                out->year  = 1969;
                out->month = 12;
                out->day   = 31;
                out->hour  = 23;
                out->min   = 59;
                out->sec   += 60;
                assert(out->sec >= 0);
            }
            out->us   = (int)extract_unit_64(&dt, 1000LL*1000*1000*1000);
            out->ps   = (int)extract_unit_64(&dt, 1000LL*1000);
            out->as   = (int)dt;
            break;

        default:
            throw std::runtime_error(
                        "NumPy datetime metadata is corrupted with invalid base unit");
    }

    return 0;
}
