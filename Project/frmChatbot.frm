VERSION 5.00
Begin VB.Form frmChatbot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Red Cross Blood Bank - Virtual Assistant (RedBot)"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtChat 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmChatbot.frx":0000
      Top             =   2880
      Width           =   12495
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Text            =   "Type your question here..."
      Top             =   7440
      Width           =   9975
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   615
      Left            =   10440
      TabIndex        =   1
      Top             =   6240
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   11280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   10440
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   9285
      Left            =   -720
      Picture         =   "frmChatbot.frx":00B2
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   14880
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   12480
      Picture         =   "frmChatbot.frx":16A04
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1065
   End
End
Attribute VB_Name = "frmChatbot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================
' FORM: frmChatbot (OFFLINE VIRTUAL ASSISTANT)
' Developer: Rachit Kumar Singh
' =========================================================



' =========================================================
' ?? REDBOT MEGA-LOGIC (SMART KEYWORD MATCHING) ??
' =========================================================
' =========================================================
' ?? REDBOT MEGA-LOGIC (SMART KEYWORD MATCHING FOR TEXTBOX) ??
' Developer: Rachit Kumar Singh
' =========================================================
Private Sub cmdSend_Click()
    Dim userMsg As String
    Dim botReply As String
    Dim lowerMsg As String
    
    ' User ne jo type kiya hai use uthana
    userMsg = Trim(txtMessage.Text)
    If userMsg = "" Then Exit Sub
    
    ' 1. USER MESSAGE (Chat box mein jodna aur agli line mein jana)
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.SelText = "You: " & userMsg & vbCrLf & vbCrLf
    
    ' Message ko chote aksharon (lowercase) mein badalna
    lowerMsg = LCase(userMsg)
    
    ' =========================================================
    ' 1. GENERAL & GREETINGS
    ' =========================================================
    If InStr(lowerMsg, "hello") > 0 Or InStr(lowerMsg, "hi ") > 0 Or lowerMsg = "hi" Or InStr(lowerMsg, "hey") > 0 Then
        botReply = "Hello! I am RedBot, your Blood Bank Assistant. How can I help you today?"
        
    ElseIf InStr(lowerMsg, "who are you") > 0 Or InStr(lowerMsg, "your name") > 0 Or InStr(lowerMsg, "what are you") > 0 Then
        botReply = "I am RedBot, an offline virtual assistant designed to guide you through this Blood Bank Management System."
        
    ' =========================================================
    ' 2. DEVELOPER & PROJECT INFO (Examiner's Favorite)
    ' =========================================================
   ElseIf InStr(lowerMsg, "developer") > 0 Or InStr(lowerMsg, "who made") > 0 Or InStr(lowerMsg, "creator") > 0 Or InStr(lowerMsg, "created by") > 0 Or InStr(lowerMsg, "project id") > 0 Then
    botReply = "This system was designed and developed by Team PRJ2531G under the BCA curriculum at Arcade Business College, Patna. Project ID: PRJ2531G."
    ElseIf InStr(lowerMsg, "college") > 0 Or InStr(lowerMsg, "university") > 0 Then
        botReply = "This project is submitted to Arcade Business College, Patna as part of the BCA curriculum."
        
    ElseIf InStr(lowerMsg, "mentor") > 0 Or InStr(lowerMsg, "guide") > 0 Or InStr(lowerMsg, "teacher") > 0 Then
        botReply = "The project mentors are Mrs. Anupama Singh and Mr. Khusdil Alam."
        
    ' =========================================================
    ' 3. DONOR MANAGEMENT
    ' =========================================================
    ElseIf (InStr(lowerMsg, "add") > 0 Or InStr(lowerMsg, "new") > 0 Or InStr(lowerMsg, "register") > 0) And InStr(lowerMsg, "donor") > 0 Then
        botReply = "To add a new donor: Go to 'Donor Management' > 'Donor Registration' in the top menu and click 'Add New'."
        
    ElseIf InStr(lowerMsg, "age") > 0 And InStr(lowerMsg, "donor") > 0 Then
        botReply = "A healthy donor must be between 18 to 65 years of age to donate blood safely."
        
    ElseIf (InStr(lowerMsg, "how often") > 0 Or InStr(lowerMsg, "how many times") > 0 Or InStr(lowerMsg, "frequency") > 0) And InStr(lowerMsg, "donate") > 0 Then
        botReply = "Males can safely donate blood every 3 months, while females can donate every 4 months. The system tracks this automatically!"
        
    ElseIf InStr(lowerMsg, "delete") > 0 And InStr(lowerMsg, "donor") > 0 Then
        botReply = "We do not hard-delete records to maintain medical history. Go to Donor Registration, search the donor, and click 'Delete' to safely Deactivate them."
        
    ' =========================================================
    ' 4. MEDICAL VITALS & RULES
    ' =========================================================
    ElseIf InStr(lowerMsg, "weight") > 0 Then
        botReply = "The minimum weight required to donate 350ml of blood is generally 45-50 Kg depending on gender. The exact limits are managed in the System Settings."
        
    ElseIf InStr(lowerMsg, "bp") > 0 Or InStr(lowerMsg, "blood pressure") > 0 Then
        botReply = "Acceptable Blood Pressure is generally between 100/60 mmHg and 150/90 mmHg. The system checks this automatically during a Visit."
        
    ElseIf InStr(lowerMsg, "hb") > 0 Or InStr(lowerMsg, "hemoglobin") > 0 Then
        botReply = "Normal Hemoglobin (HB) level should be 13.0-17.0 g/dL for Males, and 12.5-16.0 g/dL for Females."
        
    ElseIf InStr(lowerMsg, "pulse") > 0 Or InStr(lowerMsg, "heart rate") > 0 Then
        botReply = "A normal resting pulse rate should be between 60 and 100 beats per minute (bpm)."
        
    ElseIf InStr(lowerMsg, "temp") > 0 Or InStr(lowerMsg, "fever") > 0 Then
        botReply = "Body temperature must be normal (97.0 - 99.5 F). A person with a fever cannot donate blood."
        
    ElseIf InStr(lowerMsg, "unfit") > 0 Or InStr(lowerMsg, "fail") > 0 Or InStr(lowerMsg, "deferred") > 0 Then
        botReply = "If a donor's vitals (Weight, BP, HB, etc.) do not match our 'Reference Blood Range' table, the system automatically marks them as UNFIT/DEFERRED and blocks donation."
        
    ' =========================================================
    ' 5. INVENTORY & BLOOD INFO
    ' =========================================================
    ElseIf InStr(lowerMsg, "stock") > 0 Or InStr(lowerMsg, "inventory") > 0 Or InStr(lowerMsg, "available") > 0 Then
        botReply = "To check available blood bags, go to 'Inventory and Issue' > 'Blood Stock' from the main menu. You can also view the Live Analytics Dashboard!"
        
    ElseIf InStr(lowerMsg, "expire") > 0 Or InStr(lowerMsg, "expiry") > 0 Or InStr(lowerMsg, "validity") > 0 Then
        botReply = "Packed Red Blood Cells (PRBC) typically expire in 35 to 42 days. Platelets expire in just 5 days. Expired bags are marked inactive by the system."
        
    ElseIf InStr(lowerMsg, "universal") > 0 And InStr(lowerMsg, "donor") > 0 Then
        botReply = "O Negative (O-) is the universal blood donor. It can be given to patients of any blood type in emergencies."
        
    ElseIf InStr(lowerMsg, "universal") > 0 And (InStr(lowerMsg, "acceptor") > 0 Or InStr(lowerMsg, "receiver") > 0) Then
        botReply = "AB Positive (AB+) is the universal acceptor. They can receive blood from any blood type."
        
    ' =========================================================
    ' 6. PATIENT & ISSUE MODULE
    ' =========================================================
    ElseIf (InStr(lowerMsg, "issue") > 0 Or InStr(lowerMsg, "give") > 0) And InStr(lowerMsg, "blood") > 0 Then
        botReply = "To issue blood to a patient, navigate to 'Inventory and Issue' > 'Blood Issue'. Make sure Cross-Match is successful first!"
        
    ElseIf InStr(lowerMsg, "cross match") > 0 Or InStr(lowerMsg, "crossmatch") > 0 Then
        botReply = "Cross-matching is a laboratory test to ensure the donor's blood is fully compatible with the patient's blood. It is mandatory before issuing blood."
        
    ElseIf InStr(lowerMsg, "price") > 0 Or InStr(lowerMsg, "cost") > 0 Or InStr(lowerMsg, "fee") > 0 Then
        botReply = "Blood itself is freely donated, but hospitals may charge a standard processing and testing fee per bag."
        
    ' =========================================================
    ' 7. SYSTEM SETTINGS & ADMIN
    ' =========================================================
    ElseIf InStr(lowerMsg, "change limit") > 0 Or InStr(lowerMsg, "reference") > 0 Or InStr(lowerMsg, "setting") > 0 Then
        botReply = "Medical validation rules can be updated by the Admin in 'System and Settings' > 'Reference Blood Range'."
        
    ElseIf InStr(lowerMsg, "report") > 0 Or InStr(lowerMsg, "print") > 0 Then
        botReply = "You can view and print all data using the 'Report' menu. We have DataReports for Donors, Visits, and Inventory."
        
    ElseIf InStr(lowerMsg, "password") > 0 Or InStr(lowerMsg, "new user") > 0 Or InStr(lowerMsg, "staff") > 0 Then
        botReply = "Only Admins can add new users or reset passwords. Go to 'System and Settings' > 'Add New System User'."
        
    ElseIf InStr(lowerMsg, "log out") > 0 Or InStr(lowerMsg, "logout") > 0 Or InStr(lowerMsg, "exit") > 0 Then
        botReply = "To securely close your session, use 'System and Settings' > 'Log Out' or 'Exit'."
        
    ' =========================================================
    ' 8. CLOSING & FALLBACK
    ' =========================================================
    ElseIf InStr(lowerMsg, "thank") > 0 Or InStr(lowerMsg, "thx") > 0 Then
        botReply = "You are very welcome! I am always here to assist you."
        
    ElseIf InStr(lowerMsg, "good") > 0 Or InStr(lowerMsg, "great") > 0 Or InStr(lowerMsg, "awesome") > 0 Or InStr(lowerMsg, "nice") > 0 Or InStr(lowerMsg, "wow") > 0 Then
        botReply = "Glad you liked it! Rachit designed me to be as helpful and smart as possible."
        
        ' =========================================================
    ' 9. DONATION ELIGIBILITY – WHO CAN DONATE
    ' =========================================================
    
    ElseIf InStr(lowerMsg, "who can donate") > 0 Or InStr(lowerMsg, "eligible") > 0 Then
        botReply = "Any healthy person aged 18-65 years, weighing at least 50 Kg, with normal BP and Hemoglobin can donate blood."

    ElseIf InStr(lowerMsg, "can i donate") > 0 Then
        botReply = "If you are between 18-65 years old, weigh above 50 Kg, and have normal Hemoglobin and BP, you are eligible to donate."

    ElseIf InStr(lowerMsg, "diabetes") > 0 Then
        botReply = "A person with controlled diabetes may donate if they are not on insulin. Final decision depends on medical screening."

    ElseIf InStr(lowerMsg, "covid") > 0 Then
        botReply = "After recovering from COVID-19, you can donate blood after 14 days if fully symptom-free."

    ElseIf InStr(lowerMsg, "tattoo") > 0 Then
        botReply = "You can donate blood 6 months after getting a tattoo."

    ElseIf InStr(lowerMsg, "smoking") > 0 Then
        botReply = "Smokers can donate blood, but smoking should be avoided at least 2 hours before donation."

    ElseIf InStr(lowerMsg, "alcohol") > 0 Then
        botReply = "Avoid alcohol at least 24 hours before donating blood."

    ' =========================================================
    ' 10. NEXT DONATION DATE
    ' =========================================================

    ElseIf InStr(lowerMsg, "when can i donate again") > 0 Or InStr(lowerMsg, "next donation") > 0 Then
        botReply = "Males can donate again after 3 months. Females can donate again after 4 months from the last donation date."

    ElseIf InStr(lowerMsg, "how long after donation") > 0 Then
        botReply = "The minimum gap between two blood donations is 90 days for males and 120 days for females."

    ' =========================================================
    ' 11. BLOOD GROUP QUESTIONS
    ' =========================================================

    ElseIf InStr(lowerMsg, "blood group types") > 0 Then
        botReply = "There are 8 main blood groups: A+, A-, B+, B-, AB+, AB-, O+, and O-."

    ElseIf InStr(lowerMsg, "rare blood") > 0 Then
        botReply = "AB Negative (AB-) is considered one of the rarest blood groups."

    ElseIf InStr(lowerMsg, "most common blood") > 0 Then
        botReply = "O Positive (O+) is the most common blood group worldwide."

    ' =========================================================
    ' 12. DONATION PROCESS
    ' =========================================================

    ElseIf InStr(lowerMsg, "how much blood") > 0 Then
        botReply = "A standard blood donation collects approximately 350ml to 450ml of blood."

    ElseIf InStr(lowerMsg, "how long does donation take") > 0 Then
        botReply = "The blood donation process takes around 8 to 10 minutes. Including screening, the total time is about 30 minutes."

    ElseIf InStr(lowerMsg, "is blood donation safe") > 0 Then
        botReply = "Yes, blood donation is completely safe. Sterile and disposable equipment is used for every donor."

    ElseIf InStr(lowerMsg, "pain") > 0 Then
        botReply = "Blood donation may cause slight discomfort, but it is generally painless and safe."

    ' =========================================================
    ' 13. HEALTH BENEFITS
    ' =========================================================

    ElseIf InStr(lowerMsg, "benefit") > 0 Then
        botReply = "Blood donation helps save lives and also improves heart health by reducing excess iron levels."

    ElseIf InStr(lowerMsg, "why donate") > 0 Then
        botReply = "Every blood donation can save up to 3 lives. Your donation makes a real difference."

    ' =========================================================
    ' 14. SYSTEM FEATURES (Viva Booster)
    ' =========================================================

    ElseIf InStr(lowerMsg, "features") > 0 Then
        botReply = "This system includes Donor Management, Patient Management, Blood Inventory Tracking, Medical Validation, Reports, and AI-based RedBot assistance."

    ElseIf InStr(lowerMsg, "database") > 0 Then
        botReply = "The system uses a relational database to securely store donor, patient, and inventory data."

    ElseIf InStr(lowerMsg, "security") > 0 Then
        botReply = "The system uses role-based login access to ensure secure and authorized usage."

    ElseIf InStr(lowerMsg, "backup") > 0 Then
        botReply = "Admin can create a database backup from System and Settings for data safety."

    ' =========================================================
    ' 15. EMERGENCY
    ' =========================================================

    ElseIf InStr(lowerMsg, "emergency") > 0 Then
        botReply = "In emergency situations, O Negative blood can be used as a universal donor type."

    ElseIf InStr(lowerMsg, "low stock") > 0 Then
        botReply = "The system automatically highlights low stock levels in red on the dashboard."
        
    Else
    botReply = "I couldn't fully understand your question. ??" & vbCrLf & vbCrLf & _
               "You can ask me things like:" & vbCrLf & _
               "• Who can donate blood?" & vbCrLf & _
               "• When can I donate again?" & vbCrLf & _
               "• What is normal Hemoglobin (HB)?" & vbCrLf & _
               "• How to add a new donor?" & vbCrLf & _
               "• How to check blood stock?" & vbCrLf & _
               "• What is universal donor?" & vbCrLf & _
               "• Tell me about system features." & vbCrLf & vbCrLf & _
               "Try asking in simple words. I am here to help!"
End If
    
    ' 2. BOT REPLY (Chat box mein jodna, Auto-Scroll karna, aur Separator line lagana)
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.SelText = "RedBot: " & botReply & vbCrLf & "------------------------------------------------------------" & vbCrLf & vbCrLf
    
    ' 3. Typing box khali karna taaki naya message type kar sakein
    txtMessage.Text = ""
    txtMessage.SetFocus
End Sub





Private Sub Form_Load()
Text1.TabIndex = 0
End Sub

' =========================================================
' ?? SMART PLACEHOLDER LOGIC (CLICK TO CLEAR) ??
' =========================================================

' 1. Jab user TextBox par click karega (Focus aayega)
Private Sub txtMessage_GotFocus()
    ' Agar pehle se default message likha hai, toh box khali kar do
    If txtMessage.Text = "Type your question here..." Or txtMessage.Text = "Type your message here..." Then
        txtMessage.Text = ""
    End If
End Sub

' 2. Jab user bina kuch likhe bahar click kar dega (Focus hat jayega)
Private Sub txtMessage_LostFocus()
    ' Agar box ekdum khali hai, toh wapas default message likh do
    If Trim(txtMessage.Text) = "" Then
        txtMessage.Text = "Type your question here..."
    End If
End Sub

Private Sub Image3_Click()

    Dim userMsg As String
    Dim botReply As String
    Dim lowerMsg As String
    
    ' User ne jo type kiya hai use uthana
    userMsg = Trim(txtMessage.Text)
    If userMsg = "" Then Exit Sub
    
    ' 1. USER MESSAGE (Chat box mein jodna aur agli line mein jana)
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.SelText = "You: " & userMsg & vbCrLf & vbCrLf
    
    ' Message ko chote aksharon (lowercase) mein badalna
    lowerMsg = LCase(userMsg)
    
    ' =========================================================
    ' 1. GENERAL & GREETINGS
    ' =========================================================
    If InStr(lowerMsg, "hello") > 0 Or InStr(lowerMsg, "hi ") > 0 Or lowerMsg = "hi" Or InStr(lowerMsg, "hey") > 0 Then
        botReply = "Hello! I am RedBot, your Blood Bank Assistant. How can I help you today?"
        
    ElseIf InStr(lowerMsg, "who are you") > 0 Or InStr(lowerMsg, "your name") > 0 Or InStr(lowerMsg, "what are you") > 0 Then
        botReply = "I am RedBot, an offline virtual assistant designed to guide you through this Blood Bank Management System."
        
    ' =========================================================
    ' 2. DEVELOPER & PROJECT INFO (Examiner's Favorite)
    ' =========================================================
   ElseIf InStr(lowerMsg, "developer") > 0 Or InStr(lowerMsg, "who made") > 0 Or InStr(lowerMsg, "creator") > 0 Or InStr(lowerMsg, "created by") > 0 Or InStr(lowerMsg, "project id") > 0 Then
    botReply = "This system was designed and developed by Team PRJ2531G under the BCA curriculum at Arcade Business College, Patna. Project ID: PRJ2531G."
    ElseIf InStr(lowerMsg, "college") > 0 Or InStr(lowerMsg, "university") > 0 Then
        botReply = "This project is submitted to Arcade Business College, Patna as part of the BCA curriculum."
        
    ElseIf InStr(lowerMsg, "mentor") > 0 Or InStr(lowerMsg, "guide") > 0 Or InStr(lowerMsg, "teacher") > 0 Then
        botReply = "The project mentors are Mrs. Anupama Singh and Mr. Khusdil Alam."
        
    ' =========================================================
    ' 3. DONOR MANAGEMENT
    ' =========================================================
    ElseIf (InStr(lowerMsg, "add") > 0 Or InStr(lowerMsg, "new") > 0 Or InStr(lowerMsg, "register") > 0) And InStr(lowerMsg, "donor") > 0 Then
        botReply = "To add a new donor: Go to 'Donor Management' > 'Donor Registration' in the top menu and click 'Add New'."
        
    ElseIf InStr(lowerMsg, "age") > 0 And InStr(lowerMsg, "donor") > 0 Then
        botReply = "A healthy donor must be between 18 to 65 years of age to donate blood safely."
        
    ElseIf (InStr(lowerMsg, "how often") > 0 Or InStr(lowerMsg, "how many times") > 0 Or InStr(lowerMsg, "frequency") > 0) And InStr(lowerMsg, "donate") > 0 Then
        botReply = "Males can safely donate blood every 3 months, while females can donate every 4 months. The system tracks this automatically!"
        
    ElseIf InStr(lowerMsg, "delete") > 0 And InStr(lowerMsg, "donor") > 0 Then
        botReply = "We do not hard-delete records to maintain medical history. Go to Donor Registration, search the donor, and click 'Delete' to safely Deactivate them."
        
    ' =========================================================
    ' 4. MEDICAL VITALS & RULES
    ' =========================================================
    ElseIf InStr(lowerMsg, "weight") > 0 Then
        botReply = "The minimum weight required to donate 350ml of blood is generally 45-50 Kg depending on gender. The exact limits are managed in the System Settings."
        
    ElseIf InStr(lowerMsg, "bp") > 0 Or InStr(lowerMsg, "blood pressure") > 0 Then
        botReply = "Acceptable Blood Pressure is generally between 100/60 mmHg and 150/90 mmHg. The system checks this automatically during a Visit."
        
    ElseIf InStr(lowerMsg, "hb") > 0 Or InStr(lowerMsg, "hemoglobin") > 0 Then
        botReply = "Normal Hemoglobin (HB) level should be 13.0-17.0 g/dL for Males, and 12.5-16.0 g/dL for Females."
        
    ElseIf InStr(lowerMsg, "pulse") > 0 Or InStr(lowerMsg, "heart rate") > 0 Then
        botReply = "A normal resting pulse rate should be between 60 and 100 beats per minute (bpm)."
        
    ElseIf InStr(lowerMsg, "temp") > 0 Or InStr(lowerMsg, "fever") > 0 Then
        botReply = "Body temperature must be normal (97.0 - 99.5 F). A person with a fever cannot donate blood."
        
    ElseIf InStr(lowerMsg, "unfit") > 0 Or InStr(lowerMsg, "fail") > 0 Or InStr(lowerMsg, "deferred") > 0 Then
        botReply = "If a donor's vitals (Weight, BP, HB, etc.) do not match our 'Reference Blood Range' table, the system automatically marks them as UNFIT/DEFERRED and blocks donation."
        
    ' =========================================================
    ' 5. INVENTORY & BLOOD INFO
    ' =========================================================
    ElseIf InStr(lowerMsg, "stock") > 0 Or InStr(lowerMsg, "inventory") > 0 Or InStr(lowerMsg, "available") > 0 Then
        botReply = "To check available blood bags, go to 'Inventory and Issue' > 'Blood Stock' from the main menu. You can also view the Live Analytics Dashboard!"
        
    ElseIf InStr(lowerMsg, "expire") > 0 Or InStr(lowerMsg, "expiry") > 0 Or InStr(lowerMsg, "validity") > 0 Then
        botReply = "Packed Red Blood Cells (PRBC) typically expire in 35 to 42 days. Platelets expire in just 5 days. Expired bags are marked inactive by the system."
        
    ElseIf InStr(lowerMsg, "universal") > 0 And InStr(lowerMsg, "donor") > 0 Then
        botReply = "O Negative (O-) is the universal blood donor. It can be given to patients of any blood type in emergencies."
        
    ElseIf InStr(lowerMsg, "universal") > 0 And (InStr(lowerMsg, "acceptor") > 0 Or InStr(lowerMsg, "receiver") > 0) Then
        botReply = "AB Positive (AB+) is the universal acceptor. They can receive blood from any blood type."
        
    ' =========================================================
    ' 6. PATIENT & ISSUE MODULE
    ' =========================================================
    ElseIf (InStr(lowerMsg, "issue") > 0 Or InStr(lowerMsg, "give") > 0) And InStr(lowerMsg, "blood") > 0 Then
        botReply = "To issue blood to a patient, navigate to 'Inventory and Issue' > 'Blood Issue'. Make sure Cross-Match is successful first!"
        
    ElseIf InStr(lowerMsg, "cross match") > 0 Or InStr(lowerMsg, "crossmatch") > 0 Then
        botReply = "Cross-matching is a laboratory test to ensure the donor's blood is fully compatible with the patient's blood. It is mandatory before issuing blood."
        
    ElseIf InStr(lowerMsg, "price") > 0 Or InStr(lowerMsg, "cost") > 0 Or InStr(lowerMsg, "fee") > 0 Then
        botReply = "Blood itself is freely donated, but hospitals may charge a standard processing and testing fee per bag."
        
    ' =========================================================
    ' 7. SYSTEM SETTINGS & ADMIN
    ' =========================================================
    ElseIf InStr(lowerMsg, "change limit") > 0 Or InStr(lowerMsg, "reference") > 0 Or InStr(lowerMsg, "setting") > 0 Then
        botReply = "Medical validation rules can be updated by the Admin in 'System and Settings' > 'Reference Blood Range'."
        
    ElseIf InStr(lowerMsg, "report") > 0 Or InStr(lowerMsg, "print") > 0 Then
        botReply = "You can view and print all data using the 'Report' menu. We have DataReports for Donors, Visits, and Inventory."
        
    ElseIf InStr(lowerMsg, "password") > 0 Or InStr(lowerMsg, "new user") > 0 Or InStr(lowerMsg, "staff") > 0 Then
        botReply = "Only Admins can add new users or reset passwords. Go to 'System and Settings' > 'Add New System User'."
        
    ElseIf InStr(lowerMsg, "log out") > 0 Or InStr(lowerMsg, "logout") > 0 Or InStr(lowerMsg, "exit") > 0 Then
        botReply = "To securely close your session, use 'System and Settings' > 'Log Out' or 'Exit'."
        
    ' =========================================================
    ' 8. CLOSING & FALLBACK
    ' =========================================================
    ElseIf InStr(lowerMsg, "thank") > 0 Or InStr(lowerMsg, "thx") > 0 Then
        botReply = "You are very welcome! I am always here to assist you."
        
    ElseIf InStr(lowerMsg, "good") > 0 Or InStr(lowerMsg, "great") > 0 Or InStr(lowerMsg, "awesome") > 0 Or InStr(lowerMsg, "nice") > 0 Or InStr(lowerMsg, "wow") > 0 Then
        botReply = "Glad you liked it! Rachit designed me to be as helpful and smart as possible."
        
        ' =========================================================
    ' 9. DONATION ELIGIBILITY – WHO CAN DONATE
    ' =========================================================
    
    ElseIf InStr(lowerMsg, "who can donate") > 0 Or InStr(lowerMsg, "eligible") > 0 Then
        botReply = "Any healthy person aged 18-65 years, weighing at least 50 Kg, with normal BP and Hemoglobin can donate blood."

    ElseIf InStr(lowerMsg, "can i donate") > 0 Then
        botReply = "If you are between 18-65 years old, weigh above 50 Kg, and have normal Hemoglobin and BP, you are eligible to donate."

    ElseIf InStr(lowerMsg, "diabetes") > 0 Then
        botReply = "A person with controlled diabetes may donate if they are not on insulin. Final decision depends on medical screening."

    ElseIf InStr(lowerMsg, "covid") > 0 Then
        botReply = "After recovering from COVID-19, you can donate blood after 14 days if fully symptom-free."

    ElseIf InStr(lowerMsg, "tattoo") > 0 Then
        botReply = "You can donate blood 6 months after getting a tattoo."

    ElseIf InStr(lowerMsg, "smoking") > 0 Then
        botReply = "Smokers can donate blood, but smoking should be avoided at least 2 hours before donation."

    ElseIf InStr(lowerMsg, "alcohol") > 0 Then
        botReply = "Avoid alcohol at least 24 hours before donating blood."

    ' =========================================================
    ' 10. NEXT DONATION DATE
    ' =========================================================

    ElseIf InStr(lowerMsg, "when can i donate again") > 0 Or InStr(lowerMsg, "next donation") > 0 Then
        botReply = "Males can donate again after 3 months. Females can donate again after 4 months from the last donation date."

    ElseIf InStr(lowerMsg, "how long after donation") > 0 Then
        botReply = "The minimum gap between two blood donations is 90 days for males and 120 days for females."

    ' =========================================================
    ' 11. BLOOD GROUP QUESTIONS
    ' =========================================================

    ElseIf InStr(lowerMsg, "blood group types") > 0 Then
        botReply = "There are 8 main blood groups: A+, A-, B+, B-, AB+, AB-, O+, and O-."

    ElseIf InStr(lowerMsg, "rare blood") > 0 Then
        botReply = "AB Negative (AB-) is considered one of the rarest blood groups."

    ElseIf InStr(lowerMsg, "most common blood") > 0 Then
        botReply = "O Positive (O+) is the most common blood group worldwide."

    ' =========================================================
    ' 12. DONATION PROCESS
    ' =========================================================

    ElseIf InStr(lowerMsg, "how much blood") > 0 Then
        botReply = "A standard blood donation collects approximately 350ml to 450ml of blood."

    ElseIf InStr(lowerMsg, "how long does donation take") > 0 Then
        botReply = "The blood donation process takes around 8 to 10 minutes. Including screening, the total time is about 30 minutes."

    ElseIf InStr(lowerMsg, "is blood donation safe") > 0 Then
        botReply = "Yes, blood donation is completely safe. Sterile and disposable equipment is used for every donor."

    ElseIf InStr(lowerMsg, "pain") > 0 Then
        botReply = "Blood donation may cause slight discomfort, but it is generally painless and safe."

    ' =========================================================
    ' 13. HEALTH BENEFITS
    ' =========================================================

    ElseIf InStr(lowerMsg, "benefit") > 0 Then
        botReply = "Blood donation helps save lives and also improves heart health by reducing excess iron levels."

    ElseIf InStr(lowerMsg, "why donate") > 0 Then
        botReply = "Every blood donation can save up to 3 lives. Your donation makes a real difference."

    ' =========================================================
    ' 14. SYSTEM FEATURES (Viva Booster)
    ' =========================================================

    ElseIf InStr(lowerMsg, "features") > 0 Then
        botReply = "This system includes Donor Management, Patient Management, Blood Inventory Tracking, Medical Validation, Reports, and AI-based RedBot assistance."

    ElseIf InStr(lowerMsg, "database") > 0 Then
        botReply = "The system uses a relational database to securely store donor, patient, and inventory data."

    ElseIf InStr(lowerMsg, "security") > 0 Then
        botReply = "The system uses role-based login access to ensure secure and authorized usage."

    ElseIf InStr(lowerMsg, "backup") > 0 Then
        botReply = "Admin can create a database backup from System and Settings for data safety."

    ' =========================================================
    ' 15. EMERGENCY
    ' =========================================================

    ElseIf InStr(lowerMsg, "emergency") > 0 Then
        botReply = "In emergency situations, O Negative blood can be used as a universal donor type."

    ElseIf InStr(lowerMsg, "low stock") > 0 Then
        botReply = "The system automatically highlights low stock levels in red on the dashboard."
        
    Else
    botReply = "I couldn't fully understand your question. ??" & vbCrLf & vbCrLf & _
               "You can ask me things like:" & vbCrLf & _
               "• Who can donate blood?" & vbCrLf & _
               "• When can I donate again?" & vbCrLf & _
               "• What is normal Hemoglobin (HB)?" & vbCrLf & _
               "• How to add a new donor?" & vbCrLf & _
               "• How to check blood stock?" & vbCrLf & _
               "• What is universal donor?" & vbCrLf & _
               "• Tell me about system features." & vbCrLf & vbCrLf & _
               "Try asking in simple words. I am here to help!"
End If
    
    ' 2. BOT REPLY (Chat box mein jodna, Auto-Scroll karna, aur Separator line lagana)
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.SelText = "RedBot: " & botReply & vbCrLf & "------------------------------------------------------------" & vbCrLf & vbCrLf
    
    ' 3. Typing box khali karna taaki naya message type kar sakein
    txtMessage.Text = ""
    txtMessage.SetFocus


End Sub

' Enter button dabane par automatically message send ho jaye
Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSend_Click
        KeyAscii = 0
    End If
End Sub ' Enter button dabane par message send ho jaye
'Private Sub txtMessage_KeyPress(KeyAscii As Integer)
 '   If KeyAscii = 13 Then
 '       cmdSend_Click
 '     KeyAscii = 0
 '  End If
'End Sub
