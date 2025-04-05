function Get-AsciiArt {
    #getverion tag
    $ver = git describe --tags --abbrev=0 2>$null
    # Create ASCII Art for "AutoMate"
    $spaces = "                                                                                            "
    $asciiArt = @"

    o                        o                    o          o                 o                  
    <|>                      <|>                  <|\        /|>               <|>                 
    / \                      < >                  / \\o    o// \               < >                 
  o/   \o        o       o    |        o__ __o    \o/ v\  /v \o/     o__ __o/   |        o__  __o  
 <|__ __|>      <|>     <|>   o__/_   /v     v\    |   <\/>   |     /v     |    o__/_   /v      |> 
 /       \      < >     < >   |      />       <\  / \        / \   />     / \   |      />      //  
o/         \o     |       |    |      \         /  \o/        \o/   \      \o/   |      \o    o/    
/v           v\    o       o    o       o       o    |          |     o      |    o       v\  /v __o 
/>             <\   <\__ __/>    <\__    <\__ __/>   / \        / \    <\__  / \   <\__     <\/> __/> 
                                                                                                   
$spaces Version: $($ver ? "$ver" : "Unknown")                                                                                                   
                                                                                                   

"@
return $asciiArt
}