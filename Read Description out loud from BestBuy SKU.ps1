[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Add-Type –AssemblyName System.Speech
$SpeechSynthesizer = New-Object –TypeName System.Speech.Synthesis.SpeechSynthesizer

# Setup Header that works with BB's site
$BasicHeader = @{
          "user-agent"="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.2704.106 Safari/537.36 OPR/38.0.2220.41"
          "accept"="text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
          "accept-encoding"="gzip, deflate, br"
          "accept-language"="en-US,en;q=0.9"
            }

# List of skus it will run through
$SKU = @(
        "6452968"
        )

# loop through each sku
Foreach ($INDSKU in $SKU)
    {
        # Grab just the product title, and then regex out the words to speak
        $ProductTitleInvoke = Invoke-WebRequest -Uri "https://www.bestbuy.com/site/$($INDSKU).p?skuId=$($INDSKU)" -Headers $BasicHeader
        [string]$ProductTitle = ($($ProductTitleInvoke.content) | Select-string -Pattern "<title >(.*?) - Best Buy<\/title>" | % {$_.matches.groups[1].value}) -replace ("/","")
        Clear-Variable matches -ErrorAction SilentlyContinue

        # Grab just the description, and then regex out the words to speak
        $DescriptInvoke = Invoke-WebRequest -Uri "https://www.bestbuy.com/site/canopy/component/shop/product-description/v1?componentId=product-description&deviceClass=l&headerPosition=left&locale=en-US&skuId=$($INDSKU)" -Headers $BasicHeader
        [string]$Description = ($($DescriptInvoke.content) | Select-string -Pattern "<div><div>(.*?)<\/div>" | % {$_.matches.groups[1].value})
        Clear-Variable matches -ErrorAction SilentlyContinue

        # Grab just the features, and then (not done yet) regex just the bolded areas to speak
        $FeaturesInvoke = Invoke-WebRequest -uri "https://www.bestbuy.com/site/canopy/component/shop/product-features/v1?componentId=product-features&deviceClass=l&enableSpecialFeatures=true&headerPosition=left&locale=en-US&skuId=$($INDSKU)" -Headers $BasicHeader

        # Grab the Price then regex it out to be spoken
        $PriceInvoke = Invoke-WebRequest -uri "https://www.bestbuy.com/site/canopy/component/pricing/price/v1?context=buyingOptionsComponent&layout=small&skuId=$($INDSKU)&unlockedRetailOption=true" -Headers $BasicHeader
        [string]$Price = ($($PriceInvoke.content) | Select-String -Pattern "regularPrice\\\`"\:(\d+\.\d+)" | % {$_.matches.groups[1].value})
        Clear-Variable matches -ErrorAction SilentlyContinue

        # Kill this line to do full description
        $Description = $Description[0..20] -join ""

        # Gather all installed voices on computer
        $ListofVoices = $SpeechSynthesizer.GetInstalledVoices().voiceinfo

        # Select random voice of whats installed
        $SpeechSynthesizer.SelectVoice("$($ListOfVoices[(Get-Random -Minimum 0 -Maximum $($ListofVoices.count))].Name)")

        # Set speed of text
        $SpeechSynthesizer.Rate = -1  # -10 is slowest, 10 is fastest

        # Sing!
        $SpeechSynthesizer.Speak("(The $($ProductTitle) features $($Description), which is priced at $($Price) dollars)")

    }