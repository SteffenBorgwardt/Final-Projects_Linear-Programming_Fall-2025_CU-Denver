# Authors

This project is being worked on by [Brady Lamson](https://github.com/SteffenBorgwardt/Final-Projects_Linear-Programming_Fall-2025_CU-Denver/wiki/Brady-Lamson) and [Kevin Ruiz](https://github.com/SteffenBorgwardt/Final-Projects_Linear-Programming_Fall-2025_CU-Denver/wiki/Kevin-Ruiz).

# Abstract

## Project goal

This project seeks to explore the statistical technique known as $L_1$, or, LASSO regression from a linear programming perspective. This will be done by utilizing a simple toy dataset and performing LASSO regression using both the `statsmodels` library in python and by writing an AMPL model using the `amplpy` api. The goal will be to show the equivalence of the two perspectives. By doing that we can verify the correctness of our implementation and confidently explore the differences required for the two approaches.

We're pursuing this as the topic for our project as it feels like the perfect overlap between statistics and linear programming. Being able to look at tools we have used for years from a fresh perspective is exciting and we hope this project may help shed some light on regression for others in the class.

## Regression and LASSO background info

Regression at its core takes a set of points in $$n$$ dimensional space and attempts to fit a line, plane or hyperplane through that space in a way that minimizes its overall distance from those points. That is usually referred to as minimizing the residual sum of squares, or RSS. LASSO, or "Least Absolute Shrinkage and Selection Operation", is a special type of regression that includes a penalty for model complexity. It does this by shrinking some variables down to 0 while still minimizing the RSS. So this method also functions as a variable selection tool! Below is a comparison of the two using typical statistical notation.

$$
\text{Linear Regression: } \quad \min \sum_{i=1}^n (y_i - \hat{y}_i)^2
$$

$$
\text{LASSO Regression: } \quad
\min \underbrace{ \sum_{i=1}^n (y_i - \hat{y}_i)^2}_{\text{RSS}}
+
\underbrace{\lambda \sum_{j=1}^p |b_j|}_{\text{Penalty}}
$$

$$
\hat{y}_i = \sum_{j=1}^p b_j x_{ij}
$$

Here $y$ are the observed values, $\hat{y}$ are the estimated values, $b$ are the regression coefficients and $\lambda$ represents the penalty scalar. We can see the two parts clearly here. The first compares the observed values to the model’s estimated values and the second part has a weighted total of the model coefficients.

The penalty scalar dictates to what degree large coefficients are penalized and its value is decided by the user. Techniques such as cross validation are often used to find a value of $\lambda$ that minimizes the objective function.

## LASSO as a quadratic program (QP)

Before we convert from statistical notation, it is worth pointing out that the goal of regression we are discussing is minimizing the "root sum of squares". We have a square in what would be the objective function, so in reality we will be building this as a "quadratic program", or QP for short.

As for the conversion, there are some modifications that must be made. For starters, the objective function of this QP is simply the RSS portion from above. The penalty portion becomes the constraint, with some handling for the absolute value as it is not a valid operation for linear programs.

Removing that operator requires breaking up each coefficient into two parts so we let $b_j = b_j^+ - b_j^-$ where $b_j^+ , b_j^- \ge 0$. The absolute value then is just $|b_j| = b_j^+ + b_j^-$.

Therefore, the QP for LASSO regression can be written as:

$$
\min \sum_i (y_i - \hat{y}_i)^2
$$
$$
\text{subject to: } \sum_{j=1}^p (b_j^+ + b_j^-) \le t
$$

$$
\hat{y}_i = \sum_j b_j^+ x_{ij} - \sum_j b_j^- x_{ij}
$$

$$
b_j^+, b_j^- \ge 0
$$

This is the structure we will be working with for this project.

## The dataset

For this project we'll be using the classic `mtcars` toy dataset. Below is a description of the table from the R documentation.

> The data was extracted from the 1974 Motor Trend US magazine, and comprises fuel consumption and 10 aspects of automobile design and performance for 32 automobiles (1973–74 models).

The model for this dataset will be using the vehicles "miles per gallon" as the response variable and all others as predictor variables. So there will be 9 numeric predictors for just 32 rows. This is very important to note.

This table was chosen due to the small number of rows relative to the large number of variables. This is useful for this project as using too many variables on smaller datasets can cause coefficients to vary a lot and can result in many useless predictor variables. Each coefficient tries to estimate its own relationship with the response, but with so few observations, the model doesn’t have enough information to estimate all coefficients reliably. This leads to high variability and unstable coefficient estimates.

LASSO is a great option for this kind of situation as it will, by design, drag any of the unnecessary predictor coefficients to 0 and remove them from the model. We anticipate that this dataset will allow us to showcase this functionality and how that penalty will change for different values of $\lambda$ or $t$.

# Repository Navigation

TODO

# Bibliography

TODO
