# --- SMTP Configuration ---
$smtpServer = "smtp.gmail.com"
$smtpPort   = 587
$useSsl     = $true

# Your email credentials
$username = "talukder.roni.ict@gmail.com"
$password = "lasymiamgswokpgf"  # Your Gmail app password
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential ($username, $securePassword)

# --- Email Details ---
$from = $username
$to = $username             # Keep your email here for BCC trick
$bcc = @(
    "hc.dhaka@mea.gov.in",
    "visahelp.dhaka@mea.gov.in",
    "hcoffice.dhaka@mea.gov.in"
)

$subject = "Request for Obtaining Appointment Date for Double Entry Visa"
$body = @"
<html>
<body style='font-family:Segoe UI, Arial, sans-serif; font-size:14px; color:#222; line-height:1.6;'>
  <p>Dear Sir/Madam,</p>

  <p>I hope this message finds you well.</p>

  <p>
    I am writing to seek your kind assistance in obtaining an appointment date for a 
    <strong>Double Entry Visa</strong> to India. I have an appointment at the 
    <strong>Austrian Embassy in New Delhi on 5th November</strong>, and despite several attempts  including 
    submitting the required documents at the High Commission Office, I have been unable to secure a visa 
    appointment for the Double Entry Visa.
  </p>

  <p>
    Given the time-sensitive nature of my travel, I would be extremely grateful for your support in 
    facilitating an appointment at the earliest possible date.
  </p>

  <p>For your reference, I have attached the following documents:</p>

  <ul>
    <li>Visa Application Form</li>
    <li>Passport copy</li>
    <li>Appointment confirmation email</li>
    <li>Offer letter from the University Of Vienna</li>
  </ul>

  <p>Your kind consideration and prompt assistance in this matter would mean a great deal to me.</p>

  <p>Thank you very much for your time and understanding.</p>

  <br>

  <p>
    <strong>Warm regards,</strong><br>
    <strong>Md Yunus Ali Rony</strong><br>
    <strong>Passport No: A00601448</strong><br>
    <strong>Phone: 01739210312</strong><br>
    <strong>Web file No: BGDDV8AD4325</strong>
  </p>
</body>
</html>
"@


# Attachments
$attachments = @(
    "C:\Users\LENOVO\Desktop\India-visa-Daily-Mail\passport_copy.pdf"
   
)

# Send email  -Bcc $bcc
Send-MailMessage -From $from -To $to -Subject $subject -Body $body `
-SmtpServer $smtpServer -Port $smtpPort -UseSsl:$useSsl -Credential $credential -Attachments $attachments -BodyAsHtml:$true

Write-Host "Email sent successfully!"
