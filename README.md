# HelloID-Conn-Prov-Source-Cpage
HR system mostly used by Hospitals in France

| :information_source: Information |
|:---------------------------|
| This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.       |
<br />
<p align="center">
  <img src="https://tools4ever.nl/connector-logos/cpage-logo.png" width="500">
</p>

HelloID provisioning source connector for CPage On-Premises based on Oracle DB queries.
Please remember This is the first version of this connector. Any feedback will be appreciated!

<!-- TABLE OF CONTENTS -->
## Table of Contents
* [Introduction](#introduction)
* [Getting Started](#getting-started)
  * [Mappings](#mappings)
  * [Scope](#scope)
* [Setup the PowerShell connector](#setup-the-powershell-connector)
* [HelloID Docs](#helloid-docs)
* [Getting Help](#getting-help)

## Introduction
The interface to communicate with CPage On-Premises is through a set of Oracle or SQL DB queries.

For this connector we use the Oracle database and queries.

<!-- GETTING STARTED -->
## Getting Started

By using this connector you will have the ability to retrieve employee and contract data from the CPage HR system.

Connecting to Cpage is done using the full Oracle DB client. 
Please see the following page from the Oracle Docs for more information.

[Oracle Database Client Installation Guide](https://docs.oracle.com/cd/E11882_01/install.112/e47959/install.htm#NTCLI1280)

### Mappings
A basic person and contract mapping is provided. Make sure to further customize these accordingly.

### Scope

The data collection retrieved by the queries is a default set which is sufficient for HelloID to provision persons.
The queries can be changed by the customer itself to meet their requirements.

<!-- USAGE EXAMPLES -->
## Setup the PowerShell connector

1. Add a new 'Source System' to HelloID and make sure to import all the necessary files.

    - [ ] configuration.json
    - [ ] mapping.json
    - [ ] persons.ps1
    - [ ] departments.ps1

2. Fill in the required fields on the 'Configuration' tab.

<p align="left">
  <img src="https://github.com/Tools4everBV/HelloID-Conn-Prov-Source-Cpage/raw/main/assets/config.png" width="500">
</p>

## HelloID Docs
The official HelloID documentation can be found at: https://docs.helloid.com/

## Getting help

For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/hc/en-us/articles/360012557600-Configure-a-custom-PowerShell-source-system) pages

If you need help, feel free to ask questions on our [forum](https://forum.helloid.com)
