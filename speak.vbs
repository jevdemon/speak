'
' Make the computer talk
' put say.vbs somewhere in your path (like C:\WINDOWS\system32) then from a cmdline:
' Say <text to be spoken>
'
' example: c:\>Say Hello I am speaking to you!
'
'

Dim i, args, sMsg

Set args = WScript.Arguments
sMsg = vbNullstring
For i = 0 to args.Count - 1
	' sMsg = sMsg & "args(" & i & "): " & args(i) & vbcrlf
	sMsg = sMsg & args(i) & vbcrlf
Next
' WScript.Echo sMsg

' invoke SAPI to say a text string
' see http://msdn.microsoft.com/en-us/library/ms723602(v=vs.85).aspx
'

Set oVoice = CreateObject("SAPI.SpVoice") 

' the .voice property lets you switch between installed voices 
' .GetVoices gives you a list of these voice objects each of which supports
' a .GetDescription to feed a pick list.
' Ex: "Microsoft Mary" and "Microsoft Mike" and "Microsoft Sam"
' Currently only "Microsoft Sam" is installed by default and it will be the 
' default voice. See the Speech control panel to change default voice settings.
'
' example: show me the names of all installed voices

  ' Get each token in the collection returned by GetVoices
   ' For Each T In oVoice.GetVoices
   '     strVoice = T.GetDescription     'The token's name
   ' 	 wscript.echo strVoice          
   ' Next

' how to show optional dialogs if defined
' If oVoice.IsUISupported( "SpeechAudioProperties" ) Then oVoice.DisplayUI 0 , "Audio Properties" , "SpeechAudioProperties"
' If oVoice.IsUISupported( "SpeechAudioVolume") Then oVoice.DisplayUI 0 , "Audio Volume" , "SpeechAudioVolume"
' If oVoice.IsUISupported( "SpeechEngineProperties") Then oVoice.DisplayUI 0 , "Engine Properties" , "SpeechEngineProperties"


' The Rate property gets and sets the speaking rate of the voice.
' Values range from -10 to 10, the slowest to fastest speaking rates.
' http://msdn.microsoft.com/en-us/library/ms723606(v=VS.85).aspx
' rate can be specified via XML like so <rate speed="3">

oVoice.Rate = 0

' Volume property gets/sets the base volume (loudness) level of the voice.
' Values range from 0 to 100,minimum to maximum volume levels.
' http://msdn.microsoft.com/en-us/library/ms723615(v=vs.85).aspx
' volume can be specified via XML like so <volume level="100">

oVoice.Volume = 100

' the pitch can be altered in stream using XML voice markup as long as the 
' option flags don't say to ignore it. Pitch has a range of -25 to 25 
' ex: oVoice.Speak "<pitch middle='25'>" &  "Hello" , 1  


' The Speak method initiates the speaking of a text string, a text file, 
' an XML file, or a wave file by the voice.
' http://msdn.microsoft.com/en-us/library/ms723609(v=vs.85).aspx
'
' The Speak method places a text stream in the TTS engine's input queue 
' and returns a stream number. It can be called synchronously or asynchronously. 
' When called synchronously, the Speak method does not return until the text 
' has been spoken; when called asynchronously, it returns immediately, and 
' the voice speaks as a background process.


' see http://msdn.microsoft.com/en-us/library/ms720892(v=vs.85).aspx for all flags to pass Speak()
Const SVSFDefault      = 0
Const SVSFlagsAsync    = 1
Const SVSFNLPSpeakPunc = 64
Const SVSFIsFilename   = 4

' SVSFlagsAsync - Specifies that the Speak call should be asynchronous. 
'                 That is, it will return immediately after the speak request is queued.
'
' SVSFNLPSpeakPunc - Punctuation characters should be expanded into words 
'                    (e.g. "This is it." would become "This is it period").
'
' SVSFIsFilename - The string passed to the Speak method is a file name rather than text. 
'                  the string is not spoken but the file the path that points to is spoken.
'

' SpVoice.Speak( Text As String, [Flags As SpeechVoiceSpeakFlags = SVSFDefault]) As Long
'   Text - text to be spoken or path of file to be spoken if SVSFIsFilename flag is included.
'   [Optional] Flags. Default value is SVSFDefault. 
' Returns a Long variable containing the stream number. When a voice enqueues more than one 
' stream by speaking asynchronously, the stream number associates events with stream.
'
' file mode Example: oVoice.Speak "c:\mypath\mywords.txt", SVSFIsFilename

oVoice.Speak sMsg

' when voice is speaking async .WaitUntilDone(<msTimeout>) will tell you if its done
' speaking yet or if you've timed out. msTimeout is in milliseconds. if -1 there is 
' no timeout, it waits for voice to stop speaking.
' This function returns a boolean, True if voice finished speaking; False if timedout.

'If False = oVoice.WaitUntilDone(5000) Then wscript.Echo "Voice spoke for more than 5 seconds."

'
'
'
WScript.Quit 0


Sub Demo_Wav_Files()
'
' The Speakstream method initiates speaking of a sound file by the voice.
' it can also create a wav of the speech as well as play waves in the speech stream.
' http://msdn.microsoft.com/en-us/library/ms723611(v=vs.85).aspx
'

' SpeechStreamFileMode enumeration lists the access modes of a file stream.
' http://msdn.microsoft.com/en-us/library/ms720858(v=VS.85).aspx
Const SSFMOpenForRead    = 0  ' Opens an existing file as read-only.
Const SSFMOpenReadWrite  = 1  ' Opens an existing file as read-write. Not supported for wav files.
Const SSFMCreate         = 2  ' Opens existing file read-write. Else creates file then opens it read-write. Not supported for wav files.
Const SSFMCreateForWrite = 3  ' Creates file even if file exists, overwrites existing file.

' The SpeechAudioFormatType enumeration lists the supported stream formats.
' http://msdn.microsoft.com/en-us/library/ms720595(v=VS.85).aspx


    Set oVoice = CreateObject("SAPI.SpVoice") 


    ' Build a local file path and open it as a stream
    Set S = CreateObject("SAPI.SpFileStream") ' SpeechLib.SpFileStream

    Call S.Open("C:\SpeakStream.wav", SSFMCreateForWrite, False)

    ' voice speaks into the file stream and creates a WAV file
    Set oVoice.AudioOutputStream = S
    oVoice.Speak "cee : \ speak stream dot wave", SVSFNLPSpeakPunc
    S.Close

    Set oVoice = Nothing
    Set oVoice = CreateObject("SAPI.SpVoice") 

    ' voice speaks the wave file voice's stream
    Call S.Open("C:\SpeakStream.wav", , False)
    oVoice.Speak "i will now demonstrate the speak stream method."
    oVoice.SpeakStream S
    oVoice.Speak "that sounded like a wave file , but it was me."

End Sub

Sub Gort_Msg()
' Let gort speak a promo about Doug to a wave file for the OSDSC stream.

Const SSFMCreateForWrite = 3

    Set oVoice = CreateObject("SAPI.SpVoice") 
    Set S  = CreateObject("SAPI.SpFileStream") ' SpeechLib.SpFileStream
    Set S2 = CreateObject("SAPI.SpFileStream") ' SpeechLib.SpFileStream

    Call S2.Open("C:\Klaatu - Gort Stop.wav", , False)
    Call S.Open("C:\GortPromo.wav", SSFMCreateForWrite, False)
    Set oVoice.AudioOutputStream = S 
    oVoice.Rate = -2
    oVoice.Volume = 100

    sMsg = "<pitch middle='-50'>greetings this is Gort. i am so glad that deadbeat Doug got a job. "
    sMsg = sMsg & " Before moving to the Crackpot Command Center, he was sleeping on my couch in the saucer 4 weeks. "
    sMsg = sMsg & " Doug still owes me money for all the oil he drank too.  "

    oVoice.Speak sMsg

    oVoice.Speak "<pitch middle='-10'>What up My Robot?  "

    sMsg = "<pitch middle='-50'> So keep him employed so he can pay me back.  In The Morning, Adam kur-e."

    oVoice.Speak sMsg

    oVoice.SpeakStream S2  'play & record the "Gort, Deklaatu Rosco" soundbite.

    oVoice.Speak "<pitch middle='-50'>I must go now. Remember, this power can not be revoked."

    S.Close
    s2.Close

    Set oVoice = Nothing
    Set oVoice = CreateObject("SAPI.SpVoice") 

    ' Play back what we just recorded
    Call S.Open("C:\GortPromo.wav", , False)
    oVoice.Speak "<pitch middle='50'>Playback Begins..."

    Wscript.Sleep 4000

    oVoice.SpeakStream S

    Wscript.Sleep 4000

    oVoice.Speak "<pitch middle='50'>Playback Ends..."

    S.Close
    Set S = Nothing
    Set S2 = Nothing
    Set oVoice = Nothing

End Sub
