# Private Repo + Password/OTP Login + Free Live Deployment (Streamlit App)

**Goal:** Keep the code private on GitHub, allow 3 users to use the app smoothly, add password + Gmail OTP authentication, and keep everything completely free.

**Platform:** Render.com (Free Tier) — supports private GitHub repositories.

---

## Step 1: Create a Gmail App Password for Sending OTPs

You cannot use your normal Gmail password to send OTP emails. Google blocks this for security reasons.

1. Go to your Gmail account → **Manage your Google Account**.
2. Go to **Security** → Turn **2-Step Verification** ON if it is not already enabled.
3. Once 2-Step Verification is enabled, you should see the **App Passwords** option on the same page.
4. Create a new App Password. You can give it any name, such as `OTPApp`.
5. Google will generate a 16-character App Password. Copy it and store it somewhere safe. You will add this to Render later. **Do not put it anywhere in your source code.**

---

## Step 2: Create the Project Folder Structure

Create a folder on your local system with the following files:

```text
my-streamlit-app/
├── app.py
├── requirements.txt
```

### `requirements.txt`

Put the following in the file:

```text
streamlit
```

If your actual tool uses additional libraries such as `pandas`, `requests`, `numpy`, etc., add each one on a separate line.

---

## Step 3: Add Password + OTP Login Logic to `app.py`

Below is a basic working template that implements password verification followed by Gmail OTP verification.

Replace the **MAIN APP** section with your actual Streamlit application code.

```python
import streamlit as st
import smtplib
import random
from email.mime.text import MIMEText
import os

# ---------- CONFIG (Loaded from Render Environment Variables) ----------
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
APP_PASSWORD = os.environ.get("APP_PASSWORD")
RECEIVER_EMAIL = os.environ.get("RECEIVER_EMAIL")


# ---------- SESSION STATE INIT ----------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "otp" not in st.session_state:
    st.session_state.otp = None

if "otp_verified" not in st.session_state:
    st.session_state.otp_verified = False


# ---------- OTP SEND FUNCTION ----------
def send_otp(otp):
    msg = MIMEText(f"Your OTP is: {otp}")
    msg["Subject"] = "Login OTP"
    msg["From"] = GMAIL_USER
    msg["To"] = RECEIVER_EMAIL

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        server.sendmail(
            GMAIL_USER,
            RECEIVER_EMAIL,
            msg.as_string()
        )


# ---------- LOGIN SCREEN ----------
def login_screen():
    st.title("Login")

    pwd = st.text_input(
        "Password",
        type="password"
    )

    if st.button("Login"):
        if pwd == APP_PASSWORD:
            st.session_state.authenticated = True

            otp = str(random.randint(100000, 999999))
            st.session_state.otp = otp

            send_otp(otp)

            st.success(
                "Password is correct. An OTP has been sent to your email."
            )

            st.rerun()

        else:
            st.error("Incorrect password.")


# ---------- OTP SCREEN ----------
def otp_screen():
    st.title("OTP Verification")

    entered_otp = st.text_input(
        "Enter the OTP sent to your email"
    )

    if st.button("Verify"):
        if entered_otp == st.session_state.otp:
            st.session_state.otp_verified = True

            st.success(
                "Verified! Loading the app..."
            )

            st.rerun()

        else:
            st.error("Incorrect OTP.")


# ---------- FLOW CONTROL ----------
if not st.session_state.authenticated:

    login_screen()

elif not st.session_state.otp_verified:

    otp_screen()

else:

    # ================= MAIN APP =================

    st.title("My Python Tool")

    st.write(
        "Your actual Streamlit application code starts here."
    )

    # =================================================
```

**Note:** This version uses a single shared password and sends the OTP to a single email address. For a simple setup with 3 users, this can work. If each user needs their own email address and individual OTP, the authentication logic can be extended accordingly.

---

## Step 4: Push the Code to a Private GitHub Repository

Run the following commands from your project folder:

```bash
git init
git add .
git commit -m "initial commit"
git remote add origin <your-private-repository-URL>
git push -u origin main
```

When creating the repository on GitHub, make sure you select **Private**, not Public.

---

## Step 5: Create a Render Account and Connect Your Repository

1. Go to Render → sign up using GitHub.
2. Authorize Render to access your GitHub repository. You can give access only to the required repository or allow access to all repositories.
3. From the Render dashboard, click **New** → **Web Service**.
4. Select your private GitHub repository.

---

## Step 6: Configure the Render Service

Use the following settings:

| Field             | Value                                                               |
| ----------------- | ------------------------------------------------------------------- |
| **Name**          | Any name, e.g. `my-tool`                                            |
| **Environment**   | Python 3                                                            |
| **Build Command** | `pip install -r requirements.txt`                                   |
| **Start Command** | `streamlit run app.py --server.port=$PORT --server.address=0.0.0.0` |
| **Instance Type** | Free                                                                |

---

## Step 7: Add Environment Variables

Go to your Render service's **Environment** section and add the following variables:

| Key                  | Value                                             |
| -------------------- | ------------------------------------------------- |
| `GMAIL_USER`         | Your Gmail address                                |
| `GMAIL_APP_PASSWORD` | The 16-character App Password generated in Step 1 |
| `APP_PASSWORD`       | The password users will enter to access the tool  |
| `RECEIVER_EMAIL`     | The email address where the OTP should be sent    |

These values should **not** be stored in your GitHub source code. They are configured separately in Render as environment variables.

---

## Step 8: Deploy and Test

1. Click **Create Web Service**.
2. Render will pull the code from your private GitHub repository and build the application.
3. Once deployment is complete, Render will provide a live URL such as:

```text
https://my-tool.onrender.com
```

4. Share this URL with your 3 users.

Each user will need to:

**Enter password → Receive OTP → Enter OTP → Access the app**

---

# Important Things to Know

### Free-tier sleep

On Render's free tier, the service can go to sleep after a period of inactivity. When someone accesses the app again, there may be a delay while the service wakes up. This is normal behavior for a free hosting tier.

### 3 users at the same time

For a lightweight Streamlit application, 3 users should generally be a small workload. However, actual performance depends on what your application does, how much CPU/RAM it consumes, and whether it performs heavy processing.

### GitHub repository remains private

Your GitHub repository can remain **Private**. Users accessing the deployed application do not automatically get access to your GitHub source code.

### Automatic deployments

If you push a new commit to GitHub and **Auto-Deploy** is enabled in Render, Render can automatically rebuild and deploy the updated application.

---

## Overall Architecture

```text
                 ┌──────────────────────┐
                 │   Private GitHub Repo │
                 │                      │
                 │   app.py             │
                 │   requirements.txt   │
                 └──────────┬───────────┘
                            │
                            │ Deploy
                            ▼
                 ┌──────────────────────┐
                 │        Render        │
                 │                      │
                 │   Streamlit App      │
                 │   Environment Vars   │
                 └──────────┬───────────┘
                            │
                            │ HTTPS
                            ▼
                 ┌──────────────────────┐
                 │       Users (3)      │
                 │                      │
                 │ Password → OTP       │
                 └──────────┬───────────┘
                            │
                            │ OTP
                            ▼
                 ┌──────────────────────┐
                 │        Gmail         │
                 │    OTP Delivery      │
                 └──────────────────────┘
```

