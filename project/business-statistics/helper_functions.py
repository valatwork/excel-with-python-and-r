import pandas as pd
import numpy as np
import scipy as sp
import scipy.stats as stats

def generate_normal_data(mean, std_dev, size, lower_bound=None, upper_bound=None,
                         column_name='value', method='clip', decimal_places=None,
                         random_seed=None):
    """
    Generate normally distributed random data within a specified range.

    Parameters:
    mean (float): Mean of the normal distribution.
    std_dev (float): Standard deviation of the normal distribution.
    size (int or tuple): Number of data points or shape of output array.
    lower_bound (float, optional): Lower bound for the data. If None, no lower bound is applied.
    upper_bound (float, optional): Upper bound for the data. If None, no upper bound is applied.
    column_name (str, default='value'): Name of the DataFrame column.
    method (str, default='clip'): Method to handle bounds:
        - 'clip': Clip values outside the range (preserves count but distorts distribution)
        - 'truncate': Use scipy's truncated normal distribution (preserves distribution shape)
        - 'reject': Use rejection sampling (slower but preserves distribution)
    decimal_places (int, optional): Round to specified decimal places. If None, no rounding is applied.
    random_seed (int, optional): Seed for random number generator for reproducibility.

    Returns:
    pd.DataFrame: DataFrame containing the generated data.
    """
    # Input validation
    if std_dev <= 0:
        raise ValueError("Standard deviation must be positive")

    # Set random seed if provided
    if random_seed is not None:
        np.random.seed(random_seed)

    if method == 'clip':
        # Generate data and clip afterward
        data = np.random.normal(mean, std_dev, size)
        if lower_bound is not None or upper_bound is not None:
            data = np.clip(data, lower_bound, upper_bound)

    elif method == 'truncate':
        # Use truncated normal distribution
        from scipy import stats

        # Convert None bounds to appropriate values for truncnorm
        a = -np.inf if lower_bound is None else (lower_bound - mean) / std_dev
        b = np.inf if upper_bound is None else (upper_bound - mean) / std_dev

        data = stats.truncnorm.rvs(a, b, loc=mean, scale=std_dev, size=size)

    elif method == 'reject':
        # Rejection sampling method
        if isinstance(size, tuple):
            total_size = np.prod(size)
            final_shape = size
        else:
            total_size = size
            final_shape = (size,)

        data = []
        while len(data) < total_size:
            candidates = np.random.normal(mean, std_dev, total_size)
            if lower_bound is not None:
                candidates = candidates[candidates >= lower_bound]
            if upper_bound is not None:
                candidates = candidates[candidates <= upper_bound]
            data.extend(candidates[:total_size - len(data)])

        data = np.array(data[:total_size]).reshape(final_shape)

    else:
        raise ValueError("Invalid method. Choose 'clip', 'truncate', or 'reject'")

    # Round if specified
    if decimal_places is not None:
        data = np.round(data, decimal_places)

    # Create DataFrame with specified column name
    return pd.DataFrame(data, columns=[column_name])


def empirical_rule(data, lower_bound, upper_bound, column=None, inclusive='both'):
    """
    Calculate the percentage of values within a specified range.

    Parameters:
    data (pd.DataFrame or pd.Series): DataFrame or Series containing the data.
    lower_bound (float): Lower bound for the range.
    upper_bound (float): Upper bound for the range.
    column (str, optional): If data is a DataFrame, specify the column to analyze.
                            If None and data is a DataFrame, uses the first column.
    inclusive (str, default='both'): Whether bounds are inclusive.
                                     Options: 'both', 'neither', 'left', 'right'.
    Returns:
    dict: A dictionary containing the percentage and optionally the statistics.
    """
    # Handle DataFrame vs Series
    if isinstance(data, pd.DataFrame):
        if column is not None:
            series = data[column]
        else:
            series = data.iloc[:, 0]  # Use the first column if none specified
    else:
        series = data

    # Create mask based on inclusivity parameter
    if inclusive == 'both':
        mask = (series >= lower_bound) & (series <= upper_bound)
    elif inclusive == 'left':
        mask = (series >= lower_bound) & (series < upper_bound)
    elif inclusive == 'right':
        mask = (series > lower_bound) & (series <= upper_bound)
    elif inclusive == 'neither':
        mask = (series > lower_bound) & (series < upper_bound)

    # Calculate percentage
    percentage = (mask.sum() / len(series)) * 100

    return print(f'The data contains {percentage:.2f}% of values between {lower_bound} and {upper_bound}.')



def get_ci(data, confidence, method='norm'):
    """
    Calculate confidence interval for the mean.

    Parameters:
    -----------
    data : array-like, pandas Series, or DataFrame
        Input data
    confidence : float, default=0.95
        Confidence level (between 0 and 1)
    method : str, default='norm'
        Distribution to use: 'norm' for normal distribution (large samples)
                           't' for t-distribution (small samples)

    Returns:
    --------
    tuple: (lower_bound, upper_bound) of the confidence interval
    """
    # Input validation
    if not 0 < confidence < 1:
        raise ValueError("Confidence must be between 0 and 1")

    # Handle different data types
    if isinstance(data, pd.DataFrame):
        data = data.iloc[:, 0]  # Use first column if DataFrame

    # Calculate statistics
    mean = data.mean()
    if isinstance(mean, pd.Series):
        mean = mean.iloc[0]

    sem = stats.sem(data)
    if isinstance(sem, pd.Series):
        sem = sem.iloc[0]

    # Calculate critical value based on method
    if method == 't':
        df = len(data) - 1  # degrees of freedom
        critical_value = stats.t.ppf((1 + confidence) / 2, df)
    elif method == 'norm':
        critical_value = stats.norm.ppf((1 + confidence) / 2)
    else:
        raise ValueError("Method must be either 'norm' or 't'")

    # Calculate confidence interval
    margin = critical_value * sem

    upper_ci = mean + margin
    lower_ci = mean - margin

    return print(f"{lower_ci:.2f}, {upper_ci:.2f}")