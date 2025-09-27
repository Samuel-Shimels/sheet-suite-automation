function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('SME Dashboard')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// Example login function
function loginUser(email, password) {
  // Replace this with your actual authentication logic
  if(email === "admin@example.com" && password === "password123") {
    return { success: true, message: "Login successful!" };
  } else {
    return { success: false, message: "Invalid credentials." };
  }
}
