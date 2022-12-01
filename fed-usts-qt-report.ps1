
# Thanks to @Johncomiskey77 (Twitter) for extensive help with the design.

# ----------------------------------------------------------------------
function first-of-month ([DateTime] $date)
{
    Get-Date -Year $date.Year -Month $date.Month -Date 1
}
# ----------------------------------------------------------------------
# end-of-month '2022-08-01'

function end-of-month ($date)
{
    $val = Get-Date $date

    (Get-Date -Year $val.Year -Month $val.Month -Day ([DateTime]::DaysInMonth($val.Year, $val.Month))).Date
}
# ----------------------------------------------------------------------
# as-of-for-month '2022-07-01'

function as-of-for-month ($date)
{
    $end = (end-of-month $date).ToString('yyyy-MM-dd')

    $first = (first-of-month $date).ToString('yyyy-MM-dd')
   
    foreach ($elt in $result_as_of_dates.soma.asOfDates)
    {
        if ($elt -le $first)
        {
            return $elt
        }
        elseif ($elt -le $end)
        {
            $elt
        }
    }
}
# ----------------------------------------------------------------------
# maturing-in-month

function maturing-in-month ($date)
{
    $ls_as_of = as-of-for-month $date | Sort-Object
    
    $result_table = [ordered]@{}

    foreach ($elt in $ls_as_of)
    {
        $result_table[$elt] = Invoke-RestMethod ('https://markets.newyorkfed.org/api/soma/tsy/get/all/asof/{0}.json' -f $elt)
    }

    $ls_dates = $ls_as_of[1..($ls_as_of.Count - 1)]

    $ls_dates = @((first-of-month $ls_dates[-1]).ToString('yyyy-MM-dd')) + $ls_dates

    $ls_dates = $ls_dates + @((end-of-month   $ls_dates[-1]).ToString('yyyy-MM-dd'))
        
    $items = for (($i = 0), ($j = 1); $j -lt $ls_dates.Count; $i++, $j++)
    {
        # '{0} {1} {2}' -f $ls_as_of[$i], $ls_dates[$i], $ls_dates[$j]

        if ($i -eq 0)
        {
            $result_table[$ls_as_of[$i]].soma.holdings | Where-Object maturityDate -GE $ls_dates[$i] | Where-Object maturityDate -LE $ls_dates[$j]
        }
        else
        {
            $result_table[$ls_as_of[$i]].soma.holdings | Where-Object maturityDate -GT $ls_dates[$i] | Where-Object maturityDate -LE $ls_dates[$j]
        }        
    }    

    $items
}
# ----------------------------------------------------------------------
# $result_as_of_dates = Invoke-RestMethod https://markets.newyorkfed.org/api/soma/asofdates/list.json

# Remove variable each Thursday to get new as-of date
# 
# Remove-Variable -Name result_as_of_dates

$result_as_of_dates = if ($result_as_of_dates -eq $null) 
{
    Write-Host 'Retrieving SOMA as-of dates' -ForegroundColor Yellow
    Invoke-RestMethod https://markets.newyorkfed.org/api/soma/asofdates/list.json 
} 
else
{
    $result_as_of_dates
}
# ----------------------------------------------------------------------
class Summary
{
    $start
    $end

    $issued
    $maturing
    $secondary
    
    $issued_amount
    $maturing_amount
    $notes_bonds_amount
    $bills_amount
    $secondary_purchased_amount
    $secondary_sold_amount

    $inflation_compensation

    $change
}

# ----------------------------------------------------------------------

# $date = '2022-01-01'

function report-for-month ($date)
{
    Write-Host $date -ForegroundColor Yellow

    $summary = [Summary]::new()

    $summary.start = (first-of-month $date).ToString('yyyy-MM-dd')
    $summary.end   = (end-of-month   $date).ToString('yyyy-MM-dd')

    # ----------------------------------------------------------------------
    # Treasury securities purchased
    
    $summary.issued = Invoke-RestMethod ('http://www.treasurydirect.gov/TA_WS/securities/search?maturingDate={0},{1}&format=json' -f $summary.start, $summary.end)

    $summary.issued_amount = ($summary.issued | Measure-Object -Property somaAccepted -Sum).Sum

    # ----------------------------------------------------------------------
    # Secondary market operations

    $summary.secondary = Invoke-RestMethod ('https://markets.newyorkfed.org/api/tsy/all/results/details/search.json?startDate={0}&endDate={1}&securityType=both' -f $summary.start, $summary.end)
        
    $summary.secondary_purchased_amount = ($summary.secondary.treasury.auctions | Where-Object operationDirection -EQ 'P' | Measure-Object -Property totalParAmtAccepted -Sum).Sum
    
    $summary.secondary_sold_amount      = ($summary.secondary.treasury.auctions | Where-Object operationDirection -EQ 'S' | Measure-Object -Property totalParAmtAccepted -Sum).Sum
    # ----------------------------------------------------------------------
    # Treasury securities maturing

    $result_tsy = maturing-in-month $summary.start

    $summary.maturing = $result_tsy | Where-Object maturityDate -GE $summary.start | Where-Object maturityDate -LE $summary.end

    $summary.notes_bonds_amount = ($summary.maturing | Where-Object securityType -EQ NotesBonds | Measure-Object -Property parValue -Sum).Sum
    $summary.bills_amount       = ($summary.maturing | Where-Object securityType -EQ Bills      | Measure-Object -Property parValue -Sum).Sum

    $summary.maturing_amount = ($summary.maturing | Measure-Object -Property parValue -Sum).Sum
    # ----------------------------------------------------------------------
    $summary.inflation_compensation = ($summary.maturing | Measure-Object -Property inflationCompensation -Sum).Sum

    $summary.change = $summary.issued_amount + $summary.secondary_purchased_amount - $summary.secondary_sold_amount - $summary.inflation_compensation - $summary.maturing_amount

    $summary
}

$summaries = @()

# $summaries += report-for-month '2018-01-01'
# $summaries += report-for-month '2018-02-01'
# $summaries += report-for-month '2018-03-01'
# $summaries += report-for-month '2018-04-01'
# $summaries += report-for-month '2018-05-01'
# $summaries += report-for-month '2018-06-01'
# $summaries += report-for-month '2018-07-01'
# $summaries += report-for-month '2018-08-01'
# $summaries += report-for-month '2018-09-01'
# $summaries += report-for-month '2018-10-01'
# $summaries += report-for-month '2018-11-01'
# $summaries += report-for-month '2018-12-01'
# 
# $summaries += report-for-month '2019-01-01'
# $summaries += report-for-month '2019-02-01'
# $summaries += report-for-month '2019-03-01'
# $summaries += report-for-month '2019-04-01'
# $summaries += report-for-month '2019-05-01'
# $summaries += report-for-month '2019-06-01'
# $summaries += report-for-month '2019-07-01'
# $summaries += report-for-month '2019-08-01'
# $summaries += report-for-month '2019-09-01'
# $summaries += report-for-month '2019-10-01'
# $summaries += report-for-month '2019-11-01'
# $summaries += report-for-month '2019-12-01'
# 
$summaries += report-for-month '2020-01-01'
$summaries += report-for-month '2020-02-01'
$summaries += report-for-month '2020-03-01'
$summaries += report-for-month '2020-04-01'
$summaries += report-for-month '2020-05-01'
$summaries += report-for-month '2020-06-01'
$summaries += report-for-month '2020-07-01'
$summaries += report-for-month '2020-08-01'
$summaries += report-for-month '2020-09-01'
$summaries += report-for-month '2020-10-01'
$summaries += report-for-month '2020-11-01'
$summaries += report-for-month '2020-12-01'

$summaries += report-for-month '2021-01-01'
$summaries += report-for-month '2021-02-01'
$summaries += report-for-month '2021-03-01'
$summaries += report-for-month '2021-04-01'
$summaries += report-for-month '2021-05-01'
$summaries += report-for-month '2021-06-01'
$summaries += report-for-month '2021-07-01'
$summaries += report-for-month '2021-08-01'
$summaries += report-for-month '2021-09-01'
$summaries += report-for-month '2021-10-01'
$summaries += report-for-month '2021-11-01'
$summaries += report-for-month '2021-12-01'

$summaries += report-for-month '2022-01-01'
$summaries += report-for-month '2022-02-01'
$summaries += report-for-month '2022-03-01'
$summaries += report-for-month '2022-04-01'
$summaries += report-for-month '2022-05-01'
$summaries += report-for-month '2022-06-01'
$summaries += report-for-month '2022-07-01'
$summaries += report-for-month '2022-08-01'
$summaries += report-for-month '2022-09-01'
$summaries += report-for-month '2022-10-01'
$summaries += report-for-month '2022-11-01'

$summaries | Format-Table `
    start, 
    end, 
    @{ Label = 'issued_amount';          Expression = { '{0:C0}'     -f $_.issued_amount              }; Align = 'right' },
    @{ Label = 'maturing_amount';        Expression = { '{0:C0}'     -f $_.maturing_amount            }; Align = 'right' }, 
    @{ Label = 'notes_bonds';            Expression = { '{0:C0}'     -f $_.notes_bonds_amount         }; Align = 'right' }, 
    @{ Label = 'bills';                  Expression = { '{0:C0}'     -f $_.bills_amount               }; Align = 'right' }, 
    @{ Label = 'inflation_compensation'; Expression = { '{0:C0}'     -f $_.inflation_compensation     }; Align = 'right' }, 
    @{ Label = 'secondary_purchased';    Expression = { '{0:C0}'     -f $_.secondary_purchased_amount }; Align = 'right' }, 
    @{ Label = 'secondary_sold';         Expression = { '{0:C0}'     -f $_.secondary_sold_amount      }; Align = 'right' }, 
    @{ Label = 'change';                 Expression = { '{0:$#,##0}' -f $_.change                     }; Align = 'right' }

# $summaries | ConvertTo-Csv | Out-File C:\Temp\fed-usts.csv