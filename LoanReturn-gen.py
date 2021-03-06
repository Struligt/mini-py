import pandas as pd
import datetime as dt
import numpy as np


def II_eq_counts(tobin_series, num_bins):
    """ returns a pandas IndexInterval of tobin_series,
     with spacing of intervals determined so as to have equal counts of observations in each one.

    """
    num_pbin = int(len(tobin_series) / num_bins)
    obs_list = tobin_series.sort_values().tolist()
    upper_bounds = [obs_list[(i + 1) * num_pbin] for i in range(num_bins)]
    lower_bounds = [0]
    lower_bounds += upper_bounds[:-1]
    return pd.IntervalIndex.from_arrays(lower_bounds, upper_bounds), upper_bounds


def color_vals(val, threshl=[0.15, 0.30, 0.50]):
    """
    Takes a scalar and returns a string with
    the css property `'color: red'` for negative
    strings, black otherwise.
    """
    colormap = ['red', 'black', 'blue', 'green']
    color = colormap[-1]
    for i, thresh in enumerate(threshl):
        if val < thresh:
            color = colormap[i]
            break
    return 'color: %s' % color


def heatmap(pivt):
    """shows a heat map for the pivot table supplied"""
    s = pivt.style.applymap(color_vals)  # returns a style object
    with open("temp.html", "w") as f:
        f.write(s.render())


if __name__ == '__main__':

    # constants:
    pd.set_option('display.expand_frame_repr', False)
    pd.set_eng_float_format(accuracy=2, use_eng_prefix=True)
    files = ['./EFSClosedLoanSummary201611-201711.xls', './EFSClosedLoanSummary201711-201811.xls']
    removal_names = ['Unnamed', 'RIN Fees', 'PAP#', 'Carproof', 'Lien', 'Invoice']
    id_col = 'Loan'

    # 1. Readin reports
    for i, file in enumerate(files):
        if i == 0:
            df = pd.read_excel(file, 'Sheet1')
            # break
        else:
            idf = pd.read_excel(file, 'Sheet1')
            df = pd.concat([df, idf], ignore_index=True)
    print("# rows with Principal Amt = 0 :", len(df[df['Principal'] == 0]))

    # 2. Reduce
    cols_to_del = []
    for c in df.columns:
        if any(n in c for n in removal_names):
            cols_to_del.append(c)
    df.drop(cols_to_del, axis=1, inplace=True)
    df = df[df.Principal > 0]

    # 3. Convert dates
    df['Start'] = df['Start'].apply(lambda d: dt.datetime.strptime(d, '%m/%d/%Y'))
    df['End'] = df['End'].apply(lambda d: dt.datetime.strptime(d, '%m/%d/%Y'))
    df['days'] = (df['End'] - df['Start']).apply(lambda x: x.days)
    df.drop('End', axis=1, inplace=True)
    df.insert(1, 'start_month', df['Start'].apply(lambda x: x.month))

    # 4. Data clean :
    df.drop(df[df.duplicated(id_col, keep='first')].index, inplace=True)  # remove duplicates of Loan ID
    df = df.set_index([id_col])  # now make Loan ID the df index
    df['days'] = df['days'].replace(0, 1)  # avoid getting inf IRRs
    # ..... drop dealers that have little history (eg 1 loan o/s only):

    # 5a... add missing fees
    fees = {'Vinlock': 15.50, 'SystemFee': 15, 'Recovery': 15, 'Searches': 0, 'Bank charges': 10, 'LienReg': 16}

    # 5. Calculate important metrics
    internal_fees = 15.50 + 15 + 15 + 0 + 10 + 16
    df['gross_interest'] = (df['Interest'] + df['Admin Fees'] - df['Tax'] - internal_fees) / df['Principal']
    df['eff_IRR'] = df['gross_interest'] * (365 / df['days'])

    # 6. Group by Dealer ...
    go = df.groupby(['Dealer'])
    gdf = go.mean()  # df showing means
    #... add column to grouped df showing # of loans per dealer
    gdf['no.loans'] = go.count()['Principal']
    gdf = gdf[gdf['no.loans'] != 1]  # get rid of dealers with only 1 loan in period
    gdf = gdf.sort_values(by='eff_IRR', ascending=True)  # df showing means
    print(gdf.head(10), "\n")

    # 7. Pivot table, by Dealer & bucketing by principal amount along columns
    pbins, upper_b = II_eq_counts(df['Principal'], 5)
    df['PA_bin'] = pd.cut(df['Principal'], pbins)
    df_piv = df.pivot_table('eff_IRR', index=['Dealer'], columns=['PA_bin'], margins=False, fill_value=0)  # default aggfunc is mean
    df_piv = df_piv.applymap(lambda x: round(x * 100 / 1) / 100)  # pretty format
    # df_piv = df_piv.applymap(lambda x: '-' if x == 0 else x) #even prettier format
    print("Head of pivot table of effective IRRs of loans, grouped by Dealer and PA bins:")
    df_piv.columns = upper_b  # rename columns for later ease of saving
    print(df_piv.head())

    # 8. Print
    dfs_to_print = [gdf, df_piv]
    resp_names = ['2016-2018-by Dealer', '2016-2018-by Dealer vs PA']
    # ... Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('2016-2018-return-sum.xlsx', engine='xlsxwriter')
    # ...write each dataframe to a different worksheet.
    for i, d in enumerate(dfs_to_print):
        d.to_excel(writer, sheet_name=resp_names[i])
    #...commit
    writer.save()
    heatmap(df_piv)

    # 9. Exit
    print("... Done - bibi.")
