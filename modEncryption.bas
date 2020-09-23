Attribute VB_Name = "RC4"
'RC4 encryption
'Encryption and Decryption
Function encryptor(ByRef pStrDataAnda, ByRef pStrKey)

 

Dim lBytAsciiAry(255)
Dim lBytKeyAry(255)
Dim lLngIndex
Dim lBytJump
Dim lBytTemp
Dim lBytY
Dim lLngT
Dim lLngX
Dim lLngKeyLength

 

If Len(pStrKey) = 0 Then Exit Function
If Len(pStrDataAnda) = 0 Then Exit Function

        lLngKeyLength = Len(pStrKey)

        For lLngIndex = 0 To 255

        lBytKeyAry(lLngIndex) = Asc(Mid(pStrKey, ((lLngIndex) Mod (lLngKeyLength)) + 1, 1))

        Next

    For lLngIndex = 0 To 255

    lBytAsciiAry(lLngIndex) = lLngIndex

        Next

    lBytJump = 0

    For lLngIndex = 0 To 255

    lBytJump = (lBytJump + lBytAsciiAry(lLngIndex) + lBytKeyAry(lLngIndex)) Mod 256
        lBytTemp = lBytAsciiAry(lLngIndex)

        lBytAsciiAry(lLngIndex) = lBytAsciiAry(lBytJump)

        lBytAsciiAry(lBytJump) = lBytTemp

 

        Next
 

 

            lLngIndex = 0

            lBytJump = 0

            For lLngX = 1 To Len(pStrDataAnda)

                lLngIndex = (lLngIndex + 1) Mod 256

                lBytJump = (lBytJump + lBytAsciiAry(lLngIndex)) Mod 256

 

 

                lLngT = (lBytAsciiAry(lLngIndex) + lBytAsciiAry(lBytJump)) Mod 256

 

  

                lBytTemp = lBytAsciiAry(lLngIndex)

                lBytAsciiAry(lLngIndex) = lBytAsciiAry(lBytJump)

                lBytAsciiAry(lBytJump) = lBytTemp

 

                lBytY = lBytAsciiAry(lLngT)

 

      

                encryptor = encryptor & Chr(Asc(Mid(pStrDataAnda, lLngX, 1)) Xor lBytY)

            Next

 

End Function





