Function New-DynamicParameter{
    [CmdletBinding()]
    [OutputType([System.Management.Automation.RuntimeDefinedParameterDictionary])]
    Param(
        [Parameter(Mandatory)]
        [string]$Name,
        [Parameter(Mandatory)]
        [Type]$Type,
        [Parameter()]
        [switch]$Mandatory,
        [Parameter()]
        [int64]$Position,
        [Parameter()]
        $DefaultValue,
        [Parameter()]
        [string]$ParameterSetName,
        [Parameter()]
        [switch]$ValueFromPipeline,
        [Parameter()]
        [switch]$ValueFromPipelineByPropertyName,
        [Parameter()]
        [switch]$ValueFromRemainingArguments,
        [Parameter()]
        [string]$HelpMessage,
        [Parameter()]
        [array]$ValidateSet,
        [Parameter()]
        [int64[]]$ValidateLength
    )
    DynamicParam{
        IF($ValidateSet){
            $AttrCol=New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $IgnoreCaseAttr=New-Object System.Management.Automation.ParameterAttribute
            $IgnoreCaseAttr.Mandatory=$false
            $AttrCol.Clear()
            $AttrCol.Add($IgnoreCaseAttr)
            $IgnoreCaseParam=New-Object System.Management.Automation.RuntimeDefinedParameter('ValidateSetIgnoreCase',[switch],$AttrCol)
            $ParamDict=New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $ParamDict.Clear()
            $ParamDict.Add('ValidateSetIgnoreCase',$IgnoreCaseParam)
            return $ParamDict
        }

        IF($Type.Name -match "\[\]"){
            $AttrCol=New-Object System.Collections.ObjectModel.Collection[System.Attribute]
            $ValidateCountAttr=New-Object System.Management.Automation.ParameterAttribute
            $ValidateCountAttr.Mandatory=$false
            $AttrCol.Add($ValidateCountAttr)
            $ValidateCountAttr=New-Object System.Management.Automation.ValidateCountAttribute(2,2)
            $AttrCol.Add($ValidateCountAttr)
            $ValidateCountParam=New-Object System.Management.Automation.RuntimeDefinedParameter('ValidateCount',[int64[]],$AttrCol)
            $ParamDict=New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $ParamDict.Add('ValidateCount',$ValidateCountParam)
            return $ParamDict
        }
    }    

    Process{
        $AttrCol=New-Object System.Collections.ObjectModel.Collection[System.Attribute]
        IF($true){#Add Attributes to collection
            $BaseAttr=New-Object System.Management.Automation.ParameterAttribute
                IF($Mandatory){
            $BaseAttr.Mandatory=$True}ELSE{$BaseAttr.Mandatory=$False}
                IF($ValueFromPipeline){
            $Baseattr.ValueFromPipeline=$True}
                IF($Position){
            $BaseAttr.Position=$Position}
                IF($ValueFromPipelineByPropertyName){
            $BaseAttr.ValueFromPipelineByPropertyName=$True}
                IF($ValueFromRemainingArguments){
            $BaseAttr.ValueFromRemainingArguments=$True}
                IF($HelpMessage){
            $BaseAttr.HelpMessage=$HelpMessage}
                IF($ParameterSetName){
            $BaseAttr.ParameterSetName=$ParameterSetName}
            $AttrCol.Add($BaseAttr)

                IF($ValidateCount){
            $VCountAttr=New-Object System.Management.Automation.ValidateCountAttribute($ValidateCount)
            $AttrCol.Add($VCountAttr)}

                IF($ValidateSet){
            $VSetAttr=New-Object System.Management.Automation.ValidateSetAttribute($ValidateSet)
                    IF($ValidateSetIgnoreCase){
            $VSetAttr.IgnoreCase=$true}
            $AttrCol.Add($VSetAttr)}

                IF($validateLength){
            $VLengthAttr=New-Object System.Management.Automation.ValidateLengthAttribute($validateLength)
            $AttrCol.Add($VLengthAttr)}
        }#Add Attributes to collection

        IF($True){#Create Parameter, add it and the Attribute Collection to the Param Dict
            $ParamDict=New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
            $NewParam=New-Object System.Management.Automation.RuntimeDefinedParameter($Name,[system.type]$Type,$AttrCol)
                IF($DefaultValue){
            $NewParam.Value=$DefaultValue}
            $ParamDict.Add($Name,$NewParam)
        }#Create Parameter, add it and the Attribute Collection to the Param Dict
        
        return $ParamDict
    }
}