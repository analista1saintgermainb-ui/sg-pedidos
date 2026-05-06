import { Resend } from "resend";

const resend = new Resend(process.env.RESEND_API_KEY);

export default async function handler(req, res) {
  try {
    const response = await resend.emails.send({
      from: "Teste <onboarding@resend.dev>",
      to: ["SEUEMAIL@gmail.com"],
      subject: "Teste automático 🚀",
      html: "<p>Se você recebeu, deu certo!</p>"
    });

    return res.status(200).json({ success: true, response });

  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}
