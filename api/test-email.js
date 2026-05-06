import { Resend } from "resend";

export default async function handler(req, res) {
  const resend = new Resend(process.env.RESEND_API_KEY);

  try {
    await resend.emails.send({
      from: "Teste <onboarding@resend.dev>",
      to: ["SEUEMAIL@gmail.com"],
      subject: "Teste automático 🚀",
      html: "<p>Agora sim vai funcionar!</p>"
    });

    return res.status(200).json({ success: true });

  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
}
