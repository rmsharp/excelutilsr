
[![Rdoc](http://www.rdocumentation.org/badges/version/roxygen2)](http://www.rdocumentation.org/packages/roxygen2) [![Build Status](https://travis-ci.org/rmsharp/excelutilsr.svg?branch=master)](https://travis-ci.org/rmsharp/excelutilsr) [![codecov](https://codecov.io/gh/rmsharp/excelutilsr/branch/master/graph/badge.svg)](https://codecov.io/gh/rmsharp/excelutilsr)

<!-- README.md is generated from README.Rmd. Please edit that file -->
excelutilsr
===========

Introduction
------------

The goal of **excelutilsr** is to provide some basic utility functions for working with Excel workbooks and worksheets. It does not particularly contain anything novel, though, to my best knowledge it has the most complete function for easily adding styles to cells using **XLConnect**.

I find that I use **create\_wkbk**, which creates workbooks and one or more worksheets, far more than any other function, because many of my clients want output in Excel format.

I will continue to add functions to this when I develop something of general use.

Use the code as you want. Please send me any useful changes and enhancements via a pull request on *github.com*.

Installation
------------

You can install **excelutilsr** from github with:

``` r
install.packages("devtools")
devtools::install_github("rmsharp/excelutilsr")
```
