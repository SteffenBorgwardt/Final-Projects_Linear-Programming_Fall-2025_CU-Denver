from sklearn.datasets import load_iris
import statsmodels.api as sm

# Ingest data, separate out 
iris = load_iris()

X = iris.data[:, [0,1,3]]   # Features: sepal length, sepal width, petal width
y = iris.data[:, 2]         # use petal length as response

# Add a constant for the intercept
X = sm.add_constant(X)

model = sm.OLS(y, X)
# L1_wt=1 means pure L1 penalty (Lasso). alpha controls regularization strength.
results = model.fit_regularized(method='elastic_net', L1_wt=1.0, alpha=0.1)

for index, coefficient in enumerate(results.params):
    print(f"Beta {index}: {round(coefficient, 5)}")
    