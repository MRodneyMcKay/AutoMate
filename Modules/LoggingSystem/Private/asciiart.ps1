<#  
    This file is part of AutoMate.  

    AutoMate is free software: you can redistribute it and/or modify  
    it under the terms of the GNU General Public License as published by  
    the Free Software Foundation, either version 3 of the License, or  
    (at your option) any later version.  

    This program is distributed in the hope that it will be useful,  
    but WITHOUT ANY WARRANTY; without even the implied warranty of  
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the  
    GNU General Public License for more details.  

    You should have received a copy of the GNU General Public License  
    along with this program. If not, see <https://www.gnu.org/licenses/>.  
#>

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