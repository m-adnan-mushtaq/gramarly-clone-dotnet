# Build and Package Script

$ErrorActionPreference = "Stop"

Write-Host "Building Solution..."
dotnet build WordOverlayProofreader.sln -c Release

Write-Host "Publishing Overlay..."
dotnet publish src/WordOverlayProofreader.Overlay/WordOverlayProofreader.Overlay.csproj -c Release -o ./dist/overlay

Write-Host "Publishing Server..."
dotnet publish src/SuggestionServer/SuggestionServer.csproj -c Release -o ./dist/server

Write-Host "Packaging VSTO..."
# VSTO packaging is complex and usually done via MSBuild /t:Publish or Visual Studio ClickOnce.
# This is a placeholder for the VSTO publish command.
# msbuild src/WordOverlayProofreader.Addin/WordOverlayProofreader.Addin.csproj /t:Publish /p:Configuration=Release /p:PublishDir=../../dist/addin/

Write-Host "Done. Artifacts in ./dist"
