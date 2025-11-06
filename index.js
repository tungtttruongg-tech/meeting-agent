const express = require("express");
const axios = require("axios");
require("dotenv").config();

const app = express();
app.use(express.json());

app.post("/api/createMeetingAndSendEmail", async (req, res) => {
  const {
    recipientEmail,
    recipientName,
    subject,
    body,
    startTime,
    endTime
  } = req.body;

  try {
    // Lấy access token từ Azure
    const tokenRes = await axios.post(
      `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`,
      new URLSearchParams({
        grant_type: "client_credentials",
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        scope: "https://graph.microsoft.com/.default"
      })
    );

    const accessToken = tokenRes.data.access_token;

    // Tạo cuộc họp
    const meetingRes = await axios.post(
      "https://graph.microsoft.com/v1.0/me/events",
      {
        subject,
        start: { dateTime: startTime, timeZone: "SE Asia Standard Time" },
        end: { dateTime: endTime, timeZone: "SE Asia Standard Time" },
        attendees: [
          {
            emailAddress: { address: recipientEmail, name: recipientName },
            type: "required"
          }
        ],
        isOnlineMeeting: true,
        onlineMeetingProvider: "teamsForBusiness"
      },
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json"
        }
      }
    );

    const meetingLink = meetingRes.data.onlineMeeting?.joinUrl || "Không có link";

    // Gửi email
    await axios.post(
      "https://graph.microsoft.com/v1.0/me/sendMail",
      {
        message: {
          subject,
          body: {
            contentType: "HTML",
            content: `${body}<br><br><strong>Link họp:</strong> ${meetingLink}${meetingLink}</a>`
          },
          toRecipients: [
            {
              emailAddress: {
                address: recipientEmail,
                name: recipientName
              }
            }
          ]
        },
        saveToSentItems: true
      },
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json"
        }
      }
    );

    res.json({ success: true, meetingLink });
  } catch (error) {
    console.error(error.response?.data || error.message);
    res.status(500).json({ error: "Lỗi khi tạo cuộc họp hoặc gửi email." });
  }
});

app.listen(3000, () => console.log("Server đang chạy tại http://localhost:3000"));
