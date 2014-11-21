<#
.Synopsis

	Look up terminology translations and user-interface translations from actual Microsoft products.

.Link

	http://www.microsoft.com/Language/en-US/Microsoft-Terminology-API.aspx
	http://download.microsoft.com/download/1/5/D/15D3DDC6-7403-4366-BE99-AF5247ADEF1C/Microsoft-Terminology-API-SDK.pdf

.Description

	Enables user to look up terminology translations and user-interface translations from actual Microsoft products.

	Features:

	* Any-to-any language translation searches, e.g. Japanese to/from French or any other language combination.
	* Filter searches with string case and hotkey sensitivity.
	* Filter searches by product name and version.
	* Get list of languages supported by the Terminology Service API.
	* Get list of products supported by the Terminology Service API.

.Parameter Text
	This parameter is required.

	A string representing the text to translate.

.Parameter From
	This parameter is required. Example: 'en-us'

	A string representing the language code of the provided text.

	The language codes used for this parameter must be a language code returned by GetLanguages.

.Parameter To
	This parameter is required. Example: 'ru-ru'

	A string representing the language code in to which to translate the text.

	The language codes used for this parameter must be a language code returned by GetLanguages.

.Parameter Sensitivity
	This parameter is optional. Default value is "CaseInsensitive".

	A string representing the sensitivity to filter results. The value can be one of the following:

	CaseInsensitive

		Return translations in which the "From" text searched disregards the case of the text.
		A search for "Cat" would return both:
			"Cat" and "cat".

	CaseSensitive

		Return translations in which the "From" text searched takes the case of the text into account.
		Only results matching the case of the "from" are returned.
		A search for "Cat" would return:
			"Cat" but not: "cat".

	HotKeyAndCaseSensitive

		Return translations in which the "From" text searched takes the case of the text into account,
		along with any hotkeys in the string. Only results matching the case of  the "from" are returned.
		A search for "&Cat" would return:
			"&Cat" but not "&cat" or "Cat"

.Parameter Operator
	This parameter is optional. Default value is "Exact".
	A string representing the type of matching operation to use.

	The value can be one of the following:

	Exact

		Return translations in which the provided text has an exact match to the translation pair’s "From" text.

	Contains

		Return translations in which the "From" text contains the provided translation text.

	AnyWord

		Return translations in which the "From" text contains any word in the provided translation text.

		This means that a search for:
			"Lorem rutrum risus quis nulla ullamcorper"

		Can even result in the hit:
			"Lorem ipsum dolor sit amet, {0}, consectetur adipiscing elit"

		Notice that there is only one word that matches. However, realize that results with more matching words will be ranked higher.
		A one word match isn't likely to be in the top results.

.Parameter Source
	This parameter is required.
	A string representing the sources in which to search for a translation.

	The parameter must be one of the following values:

	Terms

		Microsoft terminology collections are searched.

	UiStrings

		Microsoft product strings are searched for translations.

	Both

		Microsoft terminology collections and product strings are searched.

.Parameter Unique
	This parameter is optional.

	A switch indicating whether or not only unique (that is, distinct) translations should be returned.

	If true is specified, the results are aggregated so that each distinct translation only appears once.
	If false is specified, the results are not aggregated,but each instance is returned.

.Parameter MaxTranslations
	This parameter is optional. Default value is 1.

	An integer representing the maximum number of translations to return. The maximum allowed value is 20.

.Parameter IncludeDefinitions
	This parameter is optional. Default value is "False".

	A switch indicating whether or not to include term definitions.

	If true, definitions are returned for the terms in the result set (if available in the data source).
	If unique is specified as true, the first definition for each unique set of translation pairs is used.

.Parameter Name
	This parameter is optional. Mutually exclusive with parameter "Products".

	A string representing the product for which to filter the search results.

	Valid products and versions are returned by the GetProducts switch.
	If this parameter is omitted, results are not filtered by products and versions.
	When the Name parameter is provided, the search only includes items from the UiStrings source of translations.

.Parameter Versions
	This parameter is optional. Used with parameter "Name", mutually exclusive with parameter "Products".

	A array of strings representing the product version for which to filter the search results.

	If the Versions array for a product is null or empty, results matching the product are not filtered by version.

.Parameter Products
	This parameter is optional. Mutually exclusive with parameters "Name" and "Version".
	A hashtable representing multiple products for which to filter the search results.

	Example:

		@{Windows = '7','8','8.1' ; 'Windows Server' = '2008','2012'}

	Valid products and versions are returned by the GetProducts method.

	If this parameter is omitted, results are not filtered by products and versions.
	When the Products parameter is provided, the search only includes items from the UiStrings source of translations.

.Parameter GetProducts
	The GetProducts returns the list of Microsoft products and versions for which Terminology Service API provides user-interface translations.

.Parameter GetLanguages
	The GetLanguages method returns the list of language friendly names and their codes, supported by the Terminology Service API.

.Parameter Raw
	This parameter is optional. Default value is "False".

	A switch that controls what data is returned by function.

	If this switch is not present, functon return values are:

		* Translations are returned as array of strings
		* GetProducts returns hashatble that can be directly used as the "Products" parameter value.
		* GetLanguages returns hashatble with language friendly name as key and language code as value.

	If this switch is specified, the output is a raw objects returned by the Terminology Service API.
	They provids more properties (see Microsoft Terminology API SDK PDF in links section), but the objects
	can't be used directly as parameter values.

.Example
	Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source Terms

		Description
		-----------
		Get a translation of the string 'Control Panel' from English into Russian using the Terminology Collection source.

.Example
	Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source UiStrings

		Description
		-----------
		Get a translation of the string 'Control Panel' from English into Russian using the data from actual Microsoft products UI strings.

.Example
	Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source Both -MaxTranslations 20

		Description
		-----------
		Get a translation of the string 'Control Panel' from English into Russian using the data from Terminology Collection
		source and actual Microsoft products UI strings. Returns a maximum of 20 results.

.Example
	Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source UiStrings -Name Windows -Versions '7','8','8.1' -MaxTranslations 20

		Description
		-----------
		Get a translation of the string 'Control Panel' from English into Russian using the data from the actual Microsoft products UI strings.
		Search string only in Windows versions 7, 8 and 8.1. Returns a maximum of 20 results.

.Example
	Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source UiStrings -Products @{Windows = '7','8','8.1' ; 'Windows Server' = '2008','2012'} -MaxTranslations 20

		Description
		-----------
		Get a translation of the string 'Control Panel' from English into Russian using the data from the actual Microsoft products UI strings.
		Search string only in UI strings of the Windows 7, 8 and 8.1 and Windows Server 2008 and 2012. Returns a maximum of 20 results.

.Example
	Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source UiStrings -MaxTranslations 20 -Unique -Operator AnyWord

		Description
		-----------
		Get a translation of the string 'Control Panel' from English into Russian using the data from actual Microsoft products UI strings.
		Return translations in which the "From" text contains any word in the provided translation text. Only unique (that is, distinct)
		translations are returned. Returns a maximum of 20 results.

.Example
	Get-TerminologyTranslation -Text 'Control Panel' -From 'en-us' -To 'ru-ru' -Source Both -Sensitivity CaseSensitive -MaxTranslations 20 -IncludeDefinitions -Raw

		Description
		-----------
		Get a translation of the string 'Control Panel' from English into Russian using the Terminology Collection source.
		Case-sensitive search, returns a maximum of 20 results. Include the definition of the matching term.
		Definitions are only accessible with "Raw" switch, without it only translated strings are returned.

.Example
	Get-TerminologyTranslation -GetLanguages

		Description
		-----------
		Returns the list of language friendly names and their codes, supported by the Terminology Service API.

.Example
	Get-TerminologyTranslation -GetLanguages -Raw

		Description
		-----------
		Returns the list of language codes, supported by the Terminology Service API.

.Example
	Get-TerminologyTranslation -GetProducts

		Description
		-----------
		Returns the list of Microsoft products and versions for which Terminology Service API provides user-interface translations.

.Example
	Get-TerminologyTranslation -GetProducts -Raw

		Description
		-----------
		Returns the list of Microsoft products and versions and internal their ids for which Terminology Service API provides user-interface translations.

.Outputs
	If "Raw" swicth is specified, following object are returned:

	* GetTranslations

		Type: TerminologyService.Matches

		Returns collection of Match objects. A Match object consists of the properties used to
		represent a translation pair, and to define where the translation pair comes from.

	* GetLanguages

		Type: TerminologyService.Languages

		Returns Languages collection of Language objects.

	* GetProducts

		Type: TerminologyService.Products

		Returns Products collection of Product objects.
#>
function Get-TerminologyTranslation
{
	[CmdletBinding(DefaultParameterSetName='Basic')]
	Param
	(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true,
		HelpMessage="Text to translate (e.g., Control Panel)", ParameterSetName='Basic')]
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true,
		HelpMessage="Text to translate (e.g., Control Panel)", ParameterSetName='Complex')]
		[ValidateNotNullOrEmpty()]
		[string]$Text,

		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,
		HelpMessage="Language code of the provided text (e.g., en-us)", ParameterSetName='Basic')]
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,
		HelpMessage="Language code of the provided text (e.g., en-us)", ParameterSetName='Complex')]
		[ValidateScript({[System.Globalization.Cultureinfo]::GetCultureInfo($_)})]
		[string]$From,

		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,
		HelpMessage="Language code in to which to translate the text (e.g., ru-ru)", ParameterSetName='Basic')]
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Complex')]
		[ValidateScript({[System.Globalization.Cultureinfo]::GetCultureInfo($_)})]
		[string]$To,

		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Basic')]
		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Complex')]
		[ValidateSet('CaseInsensitive', 'CaseSensitive', 'HotKeyAndCaseSensitive')]
		[string]$Sensitivity = 'CaseInsensitive',

		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Basic')]
		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Complex')]
		[ValidateSet('Exact', 'Contains', 'AnyWord')]
		[string]$Operator = 'Exact',

		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,
		HelpMessage="Sources in which to search for a translation (e.g., Terms, UiStrings or Both)",
		ParameterSetName='Basic')]
		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true, ParameterSetName='Complex')]
		[ValidateSet('Terms', 'UiStrings', 'Both')]
		[string]$Source,

		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Basic')]
		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Complex')]
		[switch]$Unique,

		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Basic')]
		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Complex')]
		[ValidateRange(1,20)]
		[int]$MaxTranslations=1,

		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Basic')]
		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Complex')]
		[switch]$IncludeDefinitions,
		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Basic')]
		[ValidateNotNullOrEmpty()]
		[string]$Name,

		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Basic')]
		[array]$Versions,

		[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true,
		HelpMessage="A hashtable representing multiple products for which to filter the search results (e.g., @{Windows = '7','8','8.1';'Windows Server' = '2008','2012'})",
		ParameterSetName='Complex')]
		[ValidateNotNullOrEmpty()]
		[hashtable]$Products,

		[Parameter(Mandatory=$false, ParameterSetName='GetProducts')]
		[switch]$GetProducts,

		[Parameter(Mandatory=$false, ParameterSetName='GetLanguages')]
		[switch]$GetLanguages,

		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Basic')]
		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='Complex')]
		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='GetProducts')]
		[Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true, ParameterSetName='GetLanguages')]
		[switch]$Raw
	)
	Begin
	{
		$Uri = 'http://api.terminology.microsoft.com/Terminology.svc?wsdl'
		Try
		{
			$tSvc = New-WebServiceProxy -Namespace 'tsvc' -Class 'tsvc' -Uri $Uri -UseDefaultCredential -ErrorAction Stop
		}
		Catch
		{
			Throw 'Can''t connect to the Microsoft Terminology Service'
		}
	}
	Process
	{
		$tsvcSensitivity = [tsvc.SearchStringComparison]::$Sensitivity
		$tsvcOperator = [tsvc.SearchOperator]::$Operator
		if($Source -eq 'Both')
		{
			$tsvcSource = [tsvc.TranslationSource]::Terms, [tsvc.TranslationSource]::UiStrings
		}
		else
		{
			$tsvcSource = [tsvc.TranslationSource]::$Source
		}
		if($Name)
		{
			$Products = @{$Name = $Versions}
		}
		if($Products)
		{
			$tsvcProducts = $Products.GetEnumerator() |
								ForEach-Object {
									$Private:tmp = New-Object tsvc.Product -Property @{Name = $_.Key}
									if($_.Value)
									{
										$_.Value |
											ForEach-Object {
												$Private:tmp.Versions += New-Object tsvc.Version -Property @{Name = $_}
											}
									}
									$Private:tmp
								}
		}
		else
		{
			$tsvcProducts = $null
		}
		if($GetProducts)
		{
			$ret = $tSvc.GetProducts()
			If(!$Raw)
			{
				[System.Collections.SortedList]$Private:tmp = @{}
				$tSvc.GetProducts() | ForEach-Object {$Private:tmp += @{$_.Name = $_.Versions.Name}}
				$ret = $Private:tmp
			}
		}
		elseif($GetLanguages)
		{
			$ret = $tSvc.GetLanguages()
			If(!$Raw)
			{
				[System.Collections.SortedList]$ret = @{}
				$tSvc.GetLanguages().Code |
							ForEach-Object {
								$Code = $_
								Try
								{
									$DisplayName = [System.Globalization.Cultureinfo]::GetCultureInfo($Code).DisplayName
								}
								Catch
								{
									$DisplayName = "Unknown ($Code)"
								}
								$ret += @{$DisplayName = $Code}
							}
			}
		}
		else
		{
			$ret = $tSvc.GetTranslations(
						$Text,
						$From,
						$To,
						$tsvcSensitivity, $true,
						$tsvcOperator, $true,
						$tsvcSource,
						$Unique, $true,
						$MaxTranslations, $true,
						$IncludeDefinitions, $true,
						$tsvcProducts
					)
			if(!$Raw)
			{
				$ret = $ret | ForEach-Object {$_.Translations | ForEach-Object {$_.TranslatedText.Trim()}}
			}
		}
		return $ret
	}
	End
	{
		$tSvc.Dispose()
	}
}