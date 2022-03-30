#!/bin/bash

dotnet build --configuration Release
dotnet pack ExcelVbaSync/ExcelVbaSync.csproj --configuration Release
