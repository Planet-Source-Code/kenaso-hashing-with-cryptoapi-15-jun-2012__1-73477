Attribute VB_Name = "modTestData"
' ***************************************************************************
' Some of the test data is the same found in FIPS 180-3 publication.
' The hashed results are accurate.
'
' REFERENCE:
'
' NIST (National Institute of Standards and Technology) Publications
' (FIPS, Special Publications)
' http://csrc.nist.gov/publications/PubsFIPS.html
'
' FIPS 180-2 (Federal Information Processing Standards Publication)
' dated 1-Aug-2002, with Change Notice 1, dated 25-Feb-2004
' http://csrc.nist.gov/publications/fips/fips180-2/FIPS180-2_changenotice.pdf
'
' FIPS 180-3 (Federal Information Processing Standards Publication)
' dated Oct-2008 (supercedes FIPS 180-2)
' http://csrc.nist.gov/publications/fips/fips180-3/fips180-3_final.pdf
'
' FIPS 180-4 (Federal Information Processing Standards Publication)
' dated Mar-2012 (supercedes FIPS 180-3)
' http://csrc.nist.gov/publications/fips/fips180-4/fips-180-4.pdf
'
' Examples of hash outputs:
' http://csrc.nist.gov/groups/ST/toolkit/examples.html
'
' Additional SHA2 information and test vectors by Aaron Gifford
'     SHA2 Information - http://www.adg.us/computers/sha.html
'     Test vectors    - http://www.adg.us/computers/sha2-1.0.zip
'
' NIST Test vectors are at http://csrc.nist.gov/cryptval/shs.htm
'
' ===========================================================================
'    DATE      NAME / eMAIL
'              DESCRIPTION
' -----------  --------------------------------------------------------------
' 24-Jun-2010  Kenneth Ives  kenaso@tx.rr.com
'              Wrote module (CryptoAPI hash)
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Public Const TEST_FILE1 As String = "Vector004.dat"   ' Excert from A. Lincoln speech
  Public Const TEST_FILE2 As String = "Vector013.dat"   ' Binary test file
  Public Const TEST_FILE3 As String = "Vector017.dat"   ' test for off-by-one
  Public Const TEST_FILE4 As String = "BigFile.dat"     ' test for large amounts of data
  
  Private Const MODULE_NAME As String = "modTestData"
  
' ***************************************************************************
' Determine the algorithm used and return pertinent information
' ***************************************************************************
Public Sub SelectResults(ByVal lngAlgorithm As enumAPI_HashAlgorithms, _
                         ByVal lngExpectedResults As Long, _
                         ByRef strTestData As String, _
                         ByRef strDataLength As String, _
                         ByRef strOutput As String)
    
    Const ROUTINE_NAME As String = "SelectResults"
    
    Select Case lngExpectedResults
           Case 0
                strTestData = "abc"
                strDataLength = "3"
           Case 1
                strTestData = "The quick brown fox jumps over the lazy dog"
                strDataLength = "43"
           Case 2
                strTestData = "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
                strDataLength = "56"
           Case 3
                strTestData = "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmn" & _
                              "oijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu"
                strDataLength = "112"
           Case 4
                strTestData = "One thousand letter 'A'"
                strDataLength = "1000"
           Case 5
                strTestData = "Excert from President Abraham Lincoln in a file named " & TEST_FILE1
                strDataLength = "1515"
           Case 6
                strTestData = "A binary file that is one byte short of 17 times size " & _
                              "of the SHA-384 and SHA-512 block lengths named " & TEST_FILE2
                strDataLength = "2175"
           Case 7
                strTestData = "The length of this binary data set is designed to test for " & _
                              "off-by-one in a file named " & TEST_FILE3
                strDataLength = "12271"
           Case 8
                strTestData = "1,000,000 letter 'a'"
                strDataLength = "1000000"
           Case 9
                strTestData = "1,000,000 binary zeroes"
                strDataLength = "1000000"
           Case Else
                InfoMsg "Cannot identify test case." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
                Exit Sub
    End Select
    
    Select Case lngAlgorithm
        Case eAPI_MD2    ' 0
             Select Case lngExpectedResults
                    Case 0: strOutput = "da853b0d3f88d99b30283a69e6ded6bb"
                    Case 1: strOutput = "03d85a0d629d2c442e987525319fc471"
                    Case 2: strOutput = "0dff6b398ad5a62ac8d97566b80c3a7f"
                    Case 3: strOutput = "2c194d0376411dc0b8485d3abe2a4b6b"
                    Case 4: strOutput = "4449e5eb12c9536a5e5d9ff2a6ce8340"
                    Case 5: strOutput = "58bacf68f1f8ed5a3515ae0607b3b511"
                    Case 6: strOutput = "bd13e2b759b4301e0b9c44a30b9579af"
                    Case 7: strOutput = "90a059bf5b87f97fea4b054761a86562"
                    Case 8: strOutput = "8c0a09ff1216ecaf95c8130953c62efd"
                    Case 9: strOutput = "0be10730b33ef0be9bc9e466cdf89fc4"
             End Select
        Case eAPI_MD4    ' 1
             Select Case lngExpectedResults
                    Case 0: strOutput = "a448017aaf21d8525fc10ae87aa6729d"
                    Case 1: strOutput = "1bee69a46ba811185c194762abaeae90"
                    Case 2: strOutput = "4691a9ec81b1a6bd1ab8557240b245c5"
                    Case 3: strOutput = "2102d1d94bd58ebf5aa25c305bb783ad"
                    Case 4: strOutput = "192fe6013889a5b6c4b2673c56ceb476"
                    Case 5: strOutput = "50385e2a1f0b9869040a289eff3abff2"
                    Case 6: strOutput = "91c8ca5afc0473a95530a45bd05aec0a"
                    Case 7: strOutput = "e62d5d4058c1363a48e6db57fbc7ae33"
                    Case 8: strOutput = "bbce80cc6bb65e5c6745e30d4eeca9a4"
                    Case 9: strOutput = "d0b30f1d5bd243c0880eab13f4c9c643"
             End Select
        Case eAPI_MD5    ' 2
             Select Case lngExpectedResults
                    Case 0: strOutput = "900150983cd24fb0d6963f7d28e17f72"
                    Case 1: strOutput = "9e107d9d372bb6826bd81d3542a419d6"
                    Case 2: strOutput = "8215ef0796a20bcaaae116d3876c664a"
                    Case 3: strOutput = "03dd8807a93175fb062dfb55dc7d359c"
                    Case 4: strOutput = "7644672d049290f0390d9c993c7d343d"
                    Case 5: strOutput = "43696c3abe0610e776cde9bf4c052421"
                    Case 6: strOutput = "ff7525c76b8c724385a040983592c91f"
                    Case 7: strOutput = "cb38900171c42fa1123a21626f657367"
                    Case 8: strOutput = "7707d6ae4e027c70eea2a935c2296f21"
                    Case 9: strOutput = "879f4bba57ed37c9ec5e5aedf9864698"
             End Select
        Case eAPI_SHA1    ' 3
             Select Case lngExpectedResults
                    Case 0: strOutput = "a9993e364706816aba3e25717850c26c9cd0d89d"
                    Case 1: strOutput = "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"
                    Case 2: strOutput = "84983e441c3bd26ebaae4aa1f95129e5e54670f1"
                    Case 3: strOutput = "a49b2446a02c645bf419f995b67091253a04a259"
                    Case 4: strOutput = "3ae3644d6777a1f56a1defeabc74af9c4b313e49"
                    Case 5: strOutput = "3728b3fd827fe2bfd0900e0586a03ffd3394e647"
                    Case 6: strOutput = "a04fdd79ddd249c71f687674329e026c57bcc378"
                    Case 7: strOutput = "3cc047623d26128cc61dfdef8af7ca473814063a"
                    Case 8: strOutput = "34aa973cd4c4daa4f61eeb2bdbad27316534016f"
                    Case 9: strOutput = "bef3595266a65a2ff36b700a75e8ed95c68210b6"
             End Select
        Case eAPI_SHA256  ' 4
             Select Case lngExpectedResults
                    Case 0: strOutput = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
                    Case 1: strOutput = "d7a8fbb307d7809469ca9abcb0082e4f8d5651e46d3cdb762d02d0bf37c9e592"
                    Case 2: strOutput = "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"
                    Case 3: strOutput = "cf5b16a778af8380036ce59e7b0492370b249b11e8f07a51afac45037afee9d1"
                    Case 4: strOutput = "c2e686823489ced2017f6059b8b239318b6364f6dcd835d0a519105a1eadd6e4"
                    Case 5: strOutput = "4d25fccf8752ce470a58cd21d90939b7eb25f3fa418dd2da4c38288ea561e600"
                    Case 6: strOutput = "8ff59c6d33c5a991088bc44dd38f037eb5ad5630c91071a221ad6943e872ac29"
                    Case 7: strOutput = "88ee6ada861083094f4c64b373657e178d88ef0a4674fce6e4e1d84e3b176afb"
                    Case 8: strOutput = "cdc76e5c9914fb9281a1c7e284d73e67f1809a48a497200e046d39ccc7112cd0"
                    Case 9: strOutput = "d29751f2649b32ff572b5e0a9f541ea660a50f94ff0beedfb0b692b924cc8025"
             End Select
        Case eAPI_SHA384  ' 5
             Select Case lngExpectedResults
                    Case 0: strOutput = "cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed8086072ba1e7cc2358baeca134c825a7"
                    Case 1: strOutput = "ca737f1014a48f4c0b6dd43cb177b0afd9e5169367544c494011e3317dbf9a509cb1e5dc1e85a941bbee3d7f2afbc9b1"
                    Case 2: strOutput = "3391fdddfc8dc7393707a65b1b4709397cf8b1d162af05abfe8f450de5f36bc6b0455a8520bc4e6f5fe95b1fe3c8452b"
                    Case 3: strOutput = "09330c33f71147e83d192fc782cd1b4753111b173b3b05d22fa08086e3b0f712fcc7c71a557e2db966c3e9fa91746039"
                    Case 4: strOutput = "7df01148677b7f18617eee3a23104f0eed6bb8c90a6046f715c9445ff43c30d69e9e7082de39c3452fd1d3afd9ba0689"
                    Case 5: strOutput = "69cc75b95280bdd9e154e743903e37b1205aa382e92e051b1f48a6db9d0203f8a17c1762d46887037275606932d3381e"
                    Case 6: strOutput = "92dca5655229b3c34796a227ff1809e273499adc2830149481224e0f54ff4483bd49834d4865e508ef53d4cd22b703ce"
                    Case 7: strOutput = "78cc6402a29eb984b8f8f888ab0102cabe7c06f0b9570e3d8d744c969db14397f58ecd14e70f324bf12d8dd4cd1ad3b2"
                    Case 8: strOutput = "9d0e1809716474cb086e834e310a4a1ced149e9c00f248527972cec5704c2a5b07b8b3dc38ecc4ebae97ddd87f3d8985"
                    Case 9: strOutput = "8a1979f9049b3fff15ea3a43a4cf84c634fd14acad1c333fecb72c588b68868b66a994386dc0cd1687b9ee2e34983b81"
             End Select
        Case eAPI_SHA512  ' 6
             Select Case lngExpectedResults
                    Case 0: strOutput = "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f"
                    Case 1: strOutput = "07e547d9586f6a73f73fbac0435ed76951218fb7d0c8d788a309d785436bbb642e93a252a954f23912547d1e8a3b5ed6e1bfd7097821233fa0538f3db854fee6"
                    Case 2: strOutput = "204a8fc6dda82f0a0ced7beb8e08a41657c16ef468b228a8279be331a703c33596fd15c13b1b07f9aa1d3bea57789ca031ad85c7a71dd70354ec631238ca3445"
                    Case 3: strOutput = "8e959b75dae313da8cf4f72814fc143f8f7779c6eb9f7fa17299aeadb6889018501d289e4900f7e4331b99dec4b5433ac7d329eeb6dd26545e96e55b874be909"
                    Case 4: strOutput = "329c52ac62d1fe731151f2b895a00475445ef74f50b979c6f7bb7cae349328c1d4cb4f7261a0ab43f936a24b000651d4a824fcdd577f211aef8f806b16afe8af"
                    Case 5: strOutput = "23450737795d2f6a13aa61adcca0df5eef6df8d8db2b42cd2ca8f783734217a73e9cabc3c9b8a8602f8aeaeb34562b6b1286846060f9809b90286b3555751f09"
                    Case 6: strOutput = "0e928db6207282bfb498ee871202f2337f4074f3a1f5055a24f08e912ac118f8101832cdb9c2f702976e629183db9bacfdd7b086c800687c3599f15de7f7b9dd"
                    Case 7: strOutput = "211bec83fbca249c53668802b857a9889428dc5120f34b3eac1603f13d1b47965c387b39ef6af15b3a44c5e7b6bbb6c1096a677dc98fc8f472737540a332f378"
                    Case 8: strOutput = "e718483d0ce769644e2e42c7bc15b4638e1f98b13b2044285632a803afa973ebde0ff244877ea60a4cb0432ce577c31beb009c5c2c49aa2e4eadb217ad8cc09b"
                    Case 9: strOutput = "ce044bc9fd43269d5bbc946cbebc3bb711341115cc4abdf2edbc3ff2c57ad4b15deb699bda257fea5aef9c6e55fcf4cf9dc25a8c3ce25f2efe90908379bff7ed"
             End Select
        Case Else
                InfoMsg "Unknown hash algorithm selected." & _
                        vbNewLine & vbNewLine & MODULE_NAME & "." & ROUTINE_NAME
    End Select
    
End Sub
