import pandas as pd


def select(df, cols={}):
    """ Select data from dataframe filtering by the columns passed to this
    function and renaming them.

    :param pandas.DataFrame df: The original dataframe.
    :param dict cols: The columns which will be selected and their respective
    names.
    :return the new dataframe.
    :rtype pandas.DataFrame.
    """

    selected = df[[k for k in cols]]
    return selected.rename(index=str, columns=cols)


def discretize(df, col, bins, labels=None):
    """ Discretize numeric columns accordingly with the number of bins passed
    as param.

    :param pandas.DataFrame df: The original dataframe.
    :param string col: The column that will be discretized.
    :param list bins: customized bins.
    :param list labels: custom labels.
    :return the new dataframe.
    :rtype pandas.DataFrame.
    """

    return pd.cut(df[col], bins, False, labels=labels)

