Overview

This repository provides a complete MATLAB workflow, used in conjunction with TreeQSM (https://github.com/InverseTampere/TreeQSM), for processing tree point cloud data and optimizing quantitative structure models (QSMs) using the Individual-tree Adaptive Optimization (IAO) method. The system comprises two main scripts:
`create_QSM.m`: Processes LAS point cloud files, generating multiple QSM models with randomized parameters.
`IAO.m`: Implements the IAO method, selecting optimal QSM parameters based on a relative RMSE comparison with reference data.
This workflow enables researchers to automate the processing of tree point clouds, generate multiple QSM models with different parameters, and select the optimal model configuration using reference data.

Installation and Requirements

MATLAB Requirements (MATLAB R2019b or later)
Download example data or prepare your own data following the structure below.

Example Data

Provided Sample Data
The repository includes 5 sample LAS files:
Tree1.las​ - Sample tree 1 point cloud data
Tree2.las​ - Sample tree 2 point cloud data
Tree3.las​ - Sample tree 3 point cloud data
Tree4.las​ - Sample tree 4 point cloud data
Tree5.las​ - Sample tree 5 point cloud data
Reference Data (referencedata.xlsx)

