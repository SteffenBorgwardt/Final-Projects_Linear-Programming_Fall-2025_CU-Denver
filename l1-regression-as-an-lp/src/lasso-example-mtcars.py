import statsmodels.api as sm
import pandas as pd

mtcars = pd.read_csv("data/mtcars.csv")
X = mtcars.drop(columns=['model', 'mpg'])
y = mtcars['mpg']

# Add a constant for the intercept
X = sm.add_constant(X)

model = sm.OLS(y, X)
# L1_wt=1 means pure L1 penalty (Lasso). alpha controls regularization strength.
results = model.fit_regularized(method='elastic_net', L1_wt=1.0, alpha=0.1)

for index, coefficient in enumerate(results.params):
    print(f"Beta {index}: {round(coefficient, 5)}")
    