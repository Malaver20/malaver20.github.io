/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./*.html",
    "./**/*.html"
  ],
  darkMode: "class",
  theme: {
    extend: {
      colors: {
        primary: "#D81B60",
        primaryHover: "#E91E63",
        primaryGlow: "rgba(216, 27, 96, 0.5)",
        secondary: "#000B18",
        surface: "#0A1128",
        surfaceLight: "rgba(255, 255, 255, 0.03)",
        textMain: "#F8FAFC",
        textMuted: "#94A3B8",
        tealPharmacy: "#0D9488",
        tealGlow: "rgba(13, 148, 136, 0.4)",
        amberGold: "#D97706"
      },
      fontFamily: {
        sans: ['Outfit', 'sans-serif'],
      },
      animation: {
        'float': 'float 6s ease-in-out infinite',
        'blob': 'blob 7s infinite',
        'glow': 'glow 3s ease-in-out infinite alternate',
      },
      keyframes: {
        float: {
          '0%, 100%': { transform: 'translateY(0)' },
          '50%': { transform: 'translateY(-20px)' },
        },
        blob: {
          '0%': { transform: 'translate(0px, 0px) scale(1)' },
          '33%': { transform: 'translate(30px, -50px) scale(1.1)' },
          '66%': { transform: 'translate(-20px, 20px) scale(0.9)' },
          '100%': { transform: 'translate(0px, 0px) scale(1)' },
        },
        glow: {
          '0%': { boxShadow: '0 0 15px rgba(216, 27, 96, 0.2)' },
          '100%': { boxShadow: '0 0 30px rgba(216, 27, 96, 0.6)' },
        }
      }
    }
  },
  plugins: [],
}
