import { useState, useRef, useEffect } from "react";
import * as THREE from "three";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";

// ══════════════════════════════════════════════════════════════════════════════
//  ⚙️  OWNER CONFIG — EDIT THESE VALUES, then redeploy
// ══════════════════════════════════════════════════════════════════════════════
const CONFIG = {
  // Your contact email (used for "Contact for Pro" and email-capture fallback)
  contactEmail: "you@example.com",

  // Your Razorpay / payment page link (create a free payment link in Razorpay dashboard).
  // Leave "" to fall back to an email contact button.
  paymentLink: "",

  // OPTIONAL: A Formspree form endpoint to collect early-access emails.
  // Sign up free at formspree.io, create a form, paste its URL here e.g.
  // "https://formspree.io/f/abcdwxyz". Leave "" to use an email (mailto) fallback.
  formspreeEndpoint: "",

  // Access codes you give to paying customers. They enter one to unlock Pro.
  // Change these to your own secret codes.
  proCodes: ["PRO-2026", "EARLYBIRD"],

  // Free-tier limit on bulk SKU rows
  freeSkuLimit: 10,

  // Pricing shown in the upgrade modal (display only)
  priceLabel: "₹999 / month",
};
// ══════════════════════════════════════════════════════════════════════════════
