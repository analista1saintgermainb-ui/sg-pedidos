import { Resend } from "resend";

export default async function handler(req, res) {
  try {
    const resend = new Resend(process.env.RESEND_API_KEY);

    const data = await resend.emails.send({
      from: "Teste <onboarding@resend.dev>",
      to: ["analista1saintgermainb@gmail.com"], // COLOQUE SEU EMAIL AQUI
      subject: "Teste funcionando 🚀",
      html: "<p>Se você recebeu isso, deu certo!</p>",
    });

    return res.status(200).json({
      success: true,
      data,
    });

  } catch (error) {
    return res.status(500).json({
      error: error.message,
    });
  }
}
