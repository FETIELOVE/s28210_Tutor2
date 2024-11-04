import pandas as pd


# Load cleaned data
def load_cleaned_data(file_path):
    return pd.read_excel(file_path)


# Function to normalize a numerical column using Min-Max scaling
def min_max_scaling(dataframe, column):
    min_val = dataframe[column].min()
    max_val = dataframe[column].max()
    dataframe[column] = (dataframe[column] - min_val) / (max_val - min_val)
    return dataframe


# Function to normalize a numerical column using Z-score normalization
def z_score_normalization(dataframe, column):
    mean = dataframe[column].mean()
    std_dev = dataframe[column].std()
    dataframe[column] = (dataframe[column] - mean) / std_dev
    return dataframe


# Function to standardize date formats to ISO format
def standardize_dates(dataframe, column):
    dataframe[column] = pd.to_datetime(dataframe[column], errors='coerce').dt.strftime('%Y-%m-%dT%H:%M:%S')
    return dataframe
