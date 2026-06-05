@{
    # PSScriptAnalyzer yapilandirmasi.
    # Asagidaki kurallar bu interaktif envanter CLI araci icin KASITLI olarak
    # devre disi birakilmistir; gerekceler belirtilmistir.
    ExcludeRules = @(
        # Arac, kullaniciya renkli ilerleme/ozet gosteren interaktif bir konsol
        # uygulamasidir; bu baglamda Write-Host dogru tercihtir.
        'PSAvoidUsingWriteHost',

        # Uzak sunuculardan veri toplarken "birden fazla yontemi sirayla dene,
        # basarisiz olani sessizce atla" deseni kullanilir (registry -> WMI vb.).
        # Bu nedenle bazi catch bloklari bilerek bostur.
        'PSAvoidUsingEmptyCatchBlock',

        # New-HtmlReport / New-DemoInventory / New-SafeCimSession gibi fonksiyonlar
        # yikici (destructive) islem yapmaz; ShouldProcess/-WhatIf gerektirmez.
        'PSUseShouldProcessForStateChangingFunctions'
    )
}
