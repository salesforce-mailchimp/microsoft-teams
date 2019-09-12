<#
.SYNOPSIS
    This function converts a wide range of input objects into array objects.
#>
function ConvertTo-Array {
    [CmdletBinding(PositionalBinding=$true)]
    param (
        [Parameter(Mandatory=$true)]
        [AllowNull()]
        $Value
    )

    # Null value
    if ($null -eq $Value) {
        return ,@()
    }

    # Type is *[]
    elseif ($Value.GetType().Name.endsWith("[]")) {
        # More than one value
        if ($Value.length -gt 1) {
            return $Value | ForEach-Object -Process { $_ }
        }

        # Exactly one value
        elseif ($Value.length -eq 1) {
            return ,@($Value[0])
        }

        # Empty array
        else {
            return ,@()
        }
    }

    # Type is List`1
    elseif ($Value.GetType().Name -eq "List``1") {
        # More than one value
        if ($Value.count -gt 1) {
            return $Value | ForEach-Object -Process { $_ }
        }

        # Exactly one value
        elseif ($Value.count -eq 1) {
            return ,@($Value[0])
        }

        # Empty array
        else {
            return ,@()
        }
    }

    # Type is any other object
    else {
        return ,@($Value)
    }
}
