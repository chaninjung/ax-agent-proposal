const { App } = require("@slack/bolt");
const axios = require("axios");

const DIFY_API_KEY = "app-424rkWacx9iO1gV3iPmDRjq0";
const DIFY_API_URL = "https://api.dify.ai/v1/chat-messages";

// Conversation tracking per user
const conversations = {};

const app = new App({
  token: process.env.SLACK_BOT_TOKEN,
  signingSecret: process.env.SLACK_SIGNING_SECRET,
  socketMode: true,
  appToken: process.env.SLACK_APP_TOKEN,
});

// Listen to messages mentioning the bot or DMs
app.event("app_mention", async ({ event, say }) => {
  await handleMessage(event.user, event.text, say);
});

app.event("message", async ({ event, say }) => {
  // Only respond to DMs (not channel messages unless mentioned)
  if (event.channel_type === "im" && !event.bot_id) {
    await handleMessage(event.user, event.text, say);
  }
});

async function handleMessage(userId, text, say) {
  try {
    // Show typing indicator
    await say("🤔 잠시만요, 처리 중입니다...");

    const convId = conversations[userId] || "";

    const response = await axios.post(
      DIFY_API_URL,
      {
        inputs: {},
        query: text,
        response_mode: "blocking",
        conversation_id: convId,
        user: userId,
      },
      {
        headers: {
          Authorization: `Bearer ${DIFY_API_KEY}`,
          "Content-Type": "application/json",
        },
        timeout: 180000,
      }
    );

    // Save conversation ID
    if (response.data.conversation_id) {
      conversations[userId] = response.data.conversation_id;
    }

    const answer = response.data.answer || "응답을 생성하지 못했습니다.";

    // Split long messages (Slack has 3000 char limit per block)
    if (answer.length > 3000) {
      const chunks = answer.match(/.{1,3000}/gs);
      for (const chunk of chunks) {
        await say(chunk);
      }
    } else {
      await say(answer);
    }
  } catch (err) {
    console.error("Error:", err.response?.data || err.message);

    // Agent mode doesn't support blocking, use streaming
    if (err.response?.data?.code === "invalid_param") {
      await handleMessageStreaming(userId, text, say);
    } else {
      await say("⚠️ 처리 중 오류가 발생했습니다. 다시 시도해주세요.");
    }
  }
}

async function handleMessageStreaming(userId, text, say) {
  try {
    const convId = conversations[userId] || "";

    const response = await axios.post(
      DIFY_API_URL,
      {
        inputs: {},
        query: text,
        response_mode: "streaming",
        conversation_id: convId,
        user: userId,
      },
      {
        headers: {
          Authorization: `Bearer ${DIFY_API_KEY}`,
          "Content-Type": "application/json",
        },
        timeout: 180000,
        responseType: "stream",
      }
    );

    let fullAnswer = "";
    let newConvId = "";

    return new Promise((resolve, reject) => {
      response.data.on("data", (chunk) => {
        const lines = chunk.toString().split("\n");
        for (const line of lines) {
          if (line.startsWith("data: ")) {
            try {
              const d = JSON.parse(line.slice(6));
              if (d.event === "agent_message") {
                fullAnswer += d.answer || "";
              }
              if (d.conversation_id) {
                newConvId = d.conversation_id;
              }
              if (d.event === "message_end") {
                // Done
              }
            } catch {}
          }
        }
      });

      response.data.on("end", async () => {
        if (newConvId) conversations[userId] = newConvId;

        if (fullAnswer) {
          if (fullAnswer.length > 3000) {
            const chunks = fullAnswer.match(/.{1,3000}/gs);
            for (const chunk of chunks) {
              await say(chunk);
            }
          } else {
            await say(fullAnswer);
          }
        } else {
          await say("응답을 생성하지 못했습니다.");
        }
        resolve();
      });

      response.data.on("error", (err) => {
        reject(err);
      });
    });
  } catch (err) {
    console.error("Streaming error:", err.message);
    await say("⚠️ 처리 중 오류가 발생했습니다.");
  }
}

// Slash command
app.command("/제안서", async ({ command, ack, say }) => {
  await ack();
  const text = command.text || "제안서를 만들어주세요";
  await handleMessage(command.user_id, text, say);
});

(async () => {
  await app.start();
  console.log("⚡️ Slack Bot is running!");
})();
