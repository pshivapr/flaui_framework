BeforeAll {
    # No setup required for locale tests
}

Describe "System Locale and Regional Settings Tests" {
    
    Context "Get-WinSystemLocale Tests" {

        It "Should have system locale set to English (United Kingdom)" {
            $systemLocale = Get-WinSystemLocale
            $systemLocale.Name | Should -Be "en-GB"
        }

        It "Should have a valid DisplayName for system locale" {
            $systemLocale = Get-WinSystemLocale
            $systemLocale.DisplayName | Should -Be "English (United Kingdom)"
            Write-Host "System Locale DisplayName: $($systemLocale.DisplayName)" -ForegroundColor Cyan
        }
    }

    Context "Get-TimeZone Tests" {
        It "Should return timezone information" {
            $timezone = Get-TimeZone
            $timezone | Should -Not -BeNullOrEmpty
        }

        It "Should have a valid timezone Id" {
            $timezone = Get-TimeZone
            $timezone.Id | Should -Be "GMT Standard Time"
            Write-Host "TimeZone Id: $($timezone.Id)" -ForegroundColor Cyan
        }

        It "Should have a valid timezone DisplayName" {
            $timezone = Get-TimeZone
            $timezone.DisplayName | Should -Be "(UTC+00:00) Dublin, Edinburgh, Lisbon, London"
            Write-Host "TimeZone DisplayName: $($timezone.DisplayName)" -ForegroundColor Cyan
        }

        It "Should have a valid StandardName" {
            $timezone = Get-TimeZone
            $timezone.StandardName | Should -Be "GMT Standard Time"
            Write-Host "TimeZone Standard Name: $($timezone.StandardName)" -ForegroundColor Cyan
        }

        It "Should have a valid DaylightName" {
            $timezone = Get-TimeZone
            $timezone.DaylightName | Should -Be "GMT Summer Time"
            Write-Host "TimeZone Daylight Name: $($timezone.DaylightName)" -ForegroundColor Cyan
        }

        It "Should have a valid base UTC offset" {
            $timezone = Get-TimeZone
            $timezone.BaseUtcOffset | Should -Be ([TimeSpan]::Zero)
            Write-Host "TimeZone UTC Offset: $($timezone.BaseUtcOffset)" -ForegroundColor Cyan
        }
    }

    Context "Get-Culture NumberFormat Tests" {
        BeforeAll {
            $culture = Get-Culture
            $numberFormat = $culture.NumberFormat
        }

        It "Should return NumberFormat information" {
            $numberFormat | Should -Not -BeNullOrEmpty
        }

        It "Should have UK currency symbol (£)" {
            # $numberFormat.CurrencySymbol | Should -Be "£"
            $numberFormat.CurrencySymbol | Should -Be ([char]0x00A3)  # 0x00A3 = £
            Write-Host "Currency Symbol: $($numberFormat.CurrencySymbol)" -ForegroundColor Cyan
        }

        It "Should have valid currency decimal separator" {
            $numberFormat.CurrencyDecimalSeparator | Should -Be "."
            Write-Host "Currency Decimal Separator: $($numberFormat.CurrencyDecimalSeparator)" -ForegroundColor Cyan
        }

        It "Should have correct decimal separator for UK (dot)" {
            $numberFormat.NumberDecimalSeparator | Should -Be "."
            Write-Host "Number Decimal Separator: $($numberFormat.NumberDecimalSeparator)" -ForegroundColor Cyan
        }

        It "Should have correct group separator for UK (comma)" {
            $numberFormat.NumberGroupSeparator | Should -Be ","
            Write-Host "Number Group Separator: $($numberFormat.NumberGroupSeparator)" -ForegroundColor Cyan
        }

        It "Should have valid currency group separator" {
            $numberFormat.CurrencyGroupSeparator | Should -Be ","
            Write-Host "Currency Group Separator: $($numberFormat.CurrencyGroupSeparator)" -ForegroundColor Cyan
        }

        It "Should have valid number of decimal digits" {
            $numberFormat.NumberDecimalDigits | Should -Be "2"
            Write-Host "Number Decimal Digits: $($numberFormat.NumberDecimalDigits)" -ForegroundColor Cyan
        }
    }

    Context "Get-Culture DateTimeFormat Tests" {
        BeforeAll {
            $culture = Get-Culture
            $dateTimeFormat = $culture.DateTimeFormat
        }

        It "Should return DateTimeFormat information" {
            $dateTimeFormat | Should -Not -BeNullOrEmpty
        }

        It "Should have a calendar type of Gregorian" {
            $dateTimeFormat.Calendar.GetType().Name | Should -Be "GregorianCalendar"
            Write-Host "Calendar Type: $($dateTimeFormat.Calendar.GetType().Name)" -ForegroundColor Cyan
        }   

        It "Should have UK short date pattern (dd/MM/yyyy)" {
            $dateTimeFormat.ShortDatePattern | Should -Be "dd/MM/yyyy"
            Write-Host "Short Date Pattern: $($dateTimeFormat.ShortDatePattern)" -ForegroundColor Cyan
        }

        It "Should have a valid long date pattern" {
            $dateTimeFormat.LongDatePattern | Should -Be "dd MMMM yyyy"
            Write-Host "Long Date Pattern: $($dateTimeFormat.LongDatePattern)" -ForegroundColor Cyan
        }

        It "Should have UK short time pattern (HH:mm)" {
            $dateTimeFormat.ShortTimePattern | Should -Be "HH:mm"
            Write-Host "Short Time Pattern: $($dateTimeFormat.ShortTimePattern)" -ForegroundColor Cyan
        }

        It "Should have a valid long time pattern" {
            $dateTimeFormat.LongTimePattern | Should -Be "HH:mm:ss"
            Write-Host "Long Time Pattern: $($dateTimeFormat.LongTimePattern)" -ForegroundColor Cyan
        }

        It "Should have a valid date separator" {
            $dateTimeFormat.DateSeparator | Should -Be "/"
            Write-Host "Date Separator: $($dateTimeFormat.DateSeparator)" -ForegroundColor Cyan
        }

        It "Should have a valid time separator" {
            $dateTimeFormat.TimeSeparator | Should -Be ":"
            Write-Host "Time Separator: $($dateTimeFormat.TimeSeparator)" -ForegroundColor Cyan
        }

        It "Should have valid AM/PM designators" {
            $dateTimeFormat.AMDesignator | Should -Be "AM"
            $dateTimeFormat.PMDesignator | Should -Be "PM"
            Write-Host "AM Designator: '$($dateTimeFormat.AMDesignator)'" -ForegroundColor Cyan
            Write-Host "PM Designator: '$($dateTimeFormat.PMDesignator)'" -ForegroundColor Cyan
        }
    }
}

AfterAll {
    Write-Host "`n=== Locale and Regional Settings Test Summary ===" -ForegroundColor Green
    Write-Host "All locale verification tests completed" -ForegroundColor Green
}
