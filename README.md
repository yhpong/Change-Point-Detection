# Change-Point-Detection
Bayesian Online Change Point Detection for 1-dimensional time series, in VBA.

Main reference is: "Bayesian Online Changepoint Detection", RP Adam, D MacKay 2007

The authors had a Matlab implementation [here](http://hips.seas.harvard.edu/content/bayesian-online-changepoint-detection),
but it does not benefit from the online capability of the algorithm, and requires a large (n_T x n_T) array in memory. It's re-written here so that data point can be fed in one by one. Default conjugate prior used is normal-inverse-gamma which is suitable for gaussian process with unknown mean and variance.
