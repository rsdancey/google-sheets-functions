This document summarizes the environment that we are using on this project

=Project Definition=
This project is designed to extract a value from a Quickbooks Enterprise Desktop installation running in multi user mode. The value is the amount on one of the company's accounts. That value is then written to a Google Sheet in a specified cell.

=The Windows Machine=
When working on the Windows Machine these are the key things that you need to know about that system:

* It is running Windows 2019 Server
* It has Quickbooks Desktop Enterprise v24, 64-bit installed
* The Quickbooks SDK (QBSDK) is v16.0 and it is installed and registered
* The Administrative user is logged in
* There is a company file, and it is open

=Quickbooks SDK=

We have extracted the QBSDK API for QBFC API calls.

They are in the file QBFC16 COM OLE Data.idl. 

This file is the SOURCE OF TRUTH for this project. When trying to determine the method signature for a COM / OLE object, reference this file. Do not make assumptions based on patterns in QBXML, or other sources like Stack Exchange or google searches. ALWAYS VERYIFY method signatures with this file before implementing code.

==Rules for Using the SDK==

These are Rules that you must follow for this project:

1. All the SDK components are registered; if you believe that a component is not registered, you are wrong and need to figure out how to find it and access it.
2. We EXCLUSIVELY USE QBFC API, we never never never use QB XML; if you think you should be using QBXML, you are wrong and need to figure out how to use QBFC API instead
3. As recently as 10pm on July 8th 2025 we had a working connection to the Quickbooks Sysetm, we triggered the authentication dialog on Quickbooks, we registered the app with Quickbooks, and we had a valid session to communicate with Quickbooks; if that level of funcationality works the problem is always with our code and not with Windows, Quickbooks or the Quickbooks SDK

==Rules for working with Windows==

In general, always use calls to the Windows API from the windows crate. There are exceptions to this rule:

1: When interacting with COM we have built our own safe interface to the VARIENT systems which uses the windowsapi crate. Do not attempt to use the varient-rs crate or the oaild crate or any other crate that you might think will help - they are fundamentally incompatible with our code and just a waste of time to attempt to integrate

2: If you find yourself trying to resolve an error related to the internal structure of a varient or VARIANT, especially if it uses the anonymous.anonymous or anonymous.anonymous.anonymous code smell, assume that you need to figure out how to address the issue by one of our safe winapi related functions, or that we need to extend our local code to handle the issue

=Working with Rust and VS.Code=

You are CoPilot, an AI coding assistant built into an IDE called VS.Code. The Terminal that you interact with is Windows Power Shell. These are rules for interacting with the Terminal

* Never use && to combine two commands - you must execute each command separately
* Building the project can take several minutes; if you execute a build command don't expect it to complete in seconds
* You don't need to change into the directory of the part of the project we're working on - I will ensure we're in the proper directory in the Terminal

After reading this document and thinking deeply about it, indicate that you accept all of these rules and will obey them by saying "I understand and will obey"