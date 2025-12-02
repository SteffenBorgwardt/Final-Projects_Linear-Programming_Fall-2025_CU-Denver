from amplpy import AMPL
from amplpy import DataFrame as AMPLDataFrame
import pandas as pd

ampl = AMPL()

df = pd.read_csv("data/mtcars.csv")
df = df.set_index('model')
adf = AMPLDataFrame.from_pandas(df)

ampl.read("models/mtcars.mod")
ampl.set_data(adf, "MODELS")
ampl.eval("display mpg;")
