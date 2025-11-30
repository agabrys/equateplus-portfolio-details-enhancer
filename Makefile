.PHONY: lint-powershell
lint-powershell:
	pwsh -Command "Invoke-ScriptAnalyzer -Path ./src -Recurse -Settings ./linter/scriptanalyzer-settings.ps1"
