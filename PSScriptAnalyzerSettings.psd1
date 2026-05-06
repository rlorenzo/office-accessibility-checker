@{
    # No global rule exclusions. Any per-finding suppression lives on the
    # specific function via [Diagnostics.CodeAnalysis.SuppressMessageAttribute]
    # so it stays local to the case it justifies.
    ExcludeRules = @()
}
