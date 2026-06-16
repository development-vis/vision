"use server";

import { Resend } from "resend";

const resend = new Resend(process.env.RESEND_API_KEY);
// TODO: fromEmail: If you verified 'vision-texas.com', use 'noreply@vision-texas.com'
// If you have NOT verified a domain yet, MUST use: 'onboarding@resend.dev'
const fromEmail = "onboarding@resend.dev"; 
const toContactEmail = "bids@vision-texas.com";
const toServiceEmail = "service@vision-texas.com";
const toCareersEmail = "careers@vision-texas.com";

export type FormState = {
  status: "success" | "error";
  message: string;
} | null;

// --- Contact Form Action ---
export async function handleContactSubmit(
  prevState: FormState,
  formData: FormData,
): Promise<FormState> {
  const firstName = formData.get("firstName");
  const lastName = formData.get("lastName");
  const email = formData.get("email") as string;
  const phone = formData.get("phone");
  const message = formData.get("message");

  try {
    await resend.emails.send({
      from: `Vision Website <${fromEmail}>`,
      to: toContactEmail,
      replyTo: email,
      subject: `Contact Inquiry: ${firstName} ${lastName}`,
      html: `
        <div style="font-family: Arial, sans-serif; color: #333;">
          <h2 style="color: #2563EB;">New Contact Inquiry</h2>
          <p><strong>Name:</strong> ${firstName} ${lastName}</p>
          <p><strong>Email:</strong> <a href="mailto:${email}">${email}</a></p>
          <p><strong>Phone:</strong> ${phone}</p>
          <hr />
          <p><strong>Message:</strong></p>
          <p style="background: #f9f9f9; padding: 15px; border-radius: 5px;">${message}</p>
        </div>
      `,
    });

    return {
      status: "success",
      message: "Your message has been sent successfully!",
    };
  } catch (error) {
    console.error("Resend Error:", error);
    return {
      status: "error",
      message: "Failed to send message. Please call us directly.",
    };
  }
}

// --- Job Application Action ---
export async function handleJobApplication(
  prevState: FormState,
  formData: FormData,
): Promise<FormState> {
  const jobTitle = formData.get("jobTitle");
  const firstName = formData.get("firstName");
  const lastName = formData.get("lastName");
  const email = formData.get("email") as string;
  const experience = formData.get("experience");
  const coverLetter = formData.get("coverLetter");
  const resume = formData.get("resume") as File;

  // Prepare Attachment
  const attachments = [];
  if (resume && resume.size > 0) {
    const buffer = Buffer.from(await resume.arrayBuffer());
    attachments.push({
      filename: resume.name,
      content: buffer,
    });
  }

  // Generate list of fields
  const detailsHtml = Array.from(formData.entries())
    .filter(([key]) => key !== "resume" && key !== "coverLetter")
    .map(([key, value]) => `<p><strong>${key}:</strong> ${value}</p>`)
    .join("");

  try {
    await resend.emails.send({
      from: `Vision Careers <${fromEmail}>`,
      to: toCareersEmail,
      replyTo: email,
      subject: `Application: ${jobTitle} - ${firstName} ${lastName}`,
      html: `
        <div style="font-family: Arial, sans-serif; color: #333;">
          <h2 style="color: #2563EB;">New Job Application</h2>
          <h3>Position: ${jobTitle}</h3>
          <p><strong>Applicant:</strong> ${firstName} ${lastName} (${experience})</p>
          <p><strong>Email:</strong> <a href="mailto:${email}">${email}</a></p>
          <hr />
          <h3>Details:</h3>
          ${detailsHtml}
          <h3>Cover Letter:</h3>
          <p>${coverLetter}</p>
        </div>
      `,
      attachments: attachments,
    });

    return {
      status: "success",
      message: "Application received! We will be in touch shortly.",
    };
  } catch (error) {
    console.error("Resend Error:", error);
    return {
      status: "error",
      message: "Application failed to send. Please try again.",
    };
  }
}

// --- Service Ticket Action ---
export async function handleServiceTicket(
  prevState: FormState,
  formData: FormData,
): Promise<FormState> {
  const ticketId = Math.floor(Math.random() * 90000) + 10000;
  const company = formData.get("companyName");
  const name = formData.get("contactName");
  const email = formData.get("email") as string;
  const urgency = formData.get("urgency");
  const systemType = formData.get("systemType");
  const message = formData.get("message");
  const attachment = formData.get("attachment") as File;

  // Prepare Attachment
  const attachments = [];
  if (attachment && attachment.size > 0) {
    const buffer = Buffer.from(await attachment.arrayBuffer());
    attachments.push({
      filename: attachment.name,
      content: buffer,
    });
  }

  const urgencyColor =
    urgency === "critical"
      ? "#DC2626"
      : urgency === "high"
        ? "#D97706"
        : "#2563EB";

  try {
    await resend.emails.send({
      from: `Service Desk <${fromEmail}>`,
      to: toServiceEmail,
      replyTo: email,
      subject: `[Ticket #${ticketId}] ${urgency?.toString().toUpperCase()} - ${company}`,
      html: `
        <div style="font-family: Arial, sans-serif; color: #333; border: 1px solid #eee; border-radius: 8px; padding: 20px;">
          <div style="border-bottom: 2px solid ${urgencyColor}; padding-bottom: 10px; margin-bottom: 20px;">
            <h2 style="margin:0; color: ${urgencyColor};">Ticket #${ticketId}</h2>
            <p style="margin:0; font-size: 14px;">Priority: ${urgency?.toString().toUpperCase()}</p>
          </div>
          
          <table style="width: 100%; border-collapse: collapse;">
            <tr><td style="padding: 5px 0; font-weight: bold;">Company:</td><td>${company}</td></tr>
            <tr><td style="padding: 5px 0; font-weight: bold;">Contact:</td><td>${name}</td></tr>
            <tr><td style="padding: 5px 0; font-weight: bold;">Email:</td><td>${email}</td></tr>
            <tr><td style="padding: 5px 0; font-weight: bold;">System:</td><td>${systemType}</td></tr>
          </table>

          <div style="background: #f3f4f6; padding: 15px; border-radius: 5px; margin-top: 20px;">
            <strong>Issue Description:</strong>
            <p style="white-space: pre-wrap;">${message}</p>
          </div>
        </div>
      `,
      attachments: attachments,
    });

    return {
      status: "success",
      message: `Ticket #${ticketId} created successfully.`,
    };
  } catch (error) {
    console.error("Resend Error:", error);
    return {
      status: "error",
      message: "Failed to create ticket. Please call support.",
    };
  }
}