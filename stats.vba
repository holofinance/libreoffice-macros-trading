REM  *****  BASIC  *****


''' Function Format: STREAK(Range, "String1", "String2")
' Returns the maximum streak of "String1" found in the given range for the given values.
' The streakis broker after a single "String2".
' 
''' Usage example: 
' Data: A1:A7="Loss","Loss","Win","Loss","Loss","Win","Win"
' STREAK(A1:A7,"Loss","Win")=2
'''
public function STREAK(vCellRangeValues as variant, val as string, negVal as string) as integer

    dim vCellValue as variant
    dim count as integer
    dim maxCount as integer
	
	' msgbox vCellValue
	
    for each vCellValue in vCellRangeValues
    	if vCellValue = val then
        	count = count + 1
    		if count > maxCount then
    			maxCount = count
    		end if
        elseif vCellValue = negVal then
        	count = 0
    	end if 
    next

    STREAK = maxCount
end function


''' Function Format: EXTSTREAK(Range, "String1", "String2")
' Returns the maximum extended streak of "String1" found in the given range for the given values.
' "Extended" means the streak continues after a single "String2" and stops at two consequent occurances of "String2".
' Often used to calculate extended losing streak aka ELS.
' 
''' Usage example: 
' Data: A1:A7="Loss","Loss","Win","Loss","Loss","Win","Win"
' EXTSTREAK(A1:A7,"Loss","Win")=3
' Comment: the 3rd Loss is canceled out by the Win before it. 
'''
public function EXTSTREAK(vCellRangeValues as variant, val as string, negVal as string) as integer

    dim vCellValue as variant
    
    dim count as integer
    dim negCount as integer
    
    dim extCount as integer
    dim maxExtCount as integer
    
	
    for each vCellValue in vCellRangeValues
    	if vCellValue = val then
        	count = count + 1
    		
    		if count = 1 then
    			if negCount <> 1 then
					extCount = extCount + 1
    			end if
    			if negCount = 1 and extCount = 0 then
    				extCount = extCount + 1
    			end if
    		else
        		extCount = extCount + 1
    		end if
    		
    		negCount = 0
    		
    		if extCount > maxExtCount then
    			maxExtCount = extCount
    		end if
    		
        elseif vCellValue = negVal then
    		negCount = negCount + 1
        	
        	count = 0
   			if negCount > 1 then
				extCount = 0
   			end if
   			
    	end if
    next

    EXTSTREAK = maxExtCount
end function
