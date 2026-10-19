/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        coninco: {
          red: "#E51A23",
          dark: "#A70212",
          ink: "#3E3E3F",
          muted: "#7D7F82",
          line: "#E7E8EA",
          soft: "#FFF2F3",
        },
      },
      boxShadow: {
        soft: "0 18px 50px rgba(62, 62, 63, 0.08)",
        red: "0 14px 32px rgba(229, 26, 35, 0.22)",
      },
      keyframes: {
        float: { "0%, 100%": { transform: "translateY(0)" }, "50%": { transform: "translateY(-12px)" } },
        shimmer: { "0%": { backgroundPosition: "200% 0" }, "100%": { backgroundPosition: "-200% 0" } },
      },
      animation: {
        float: "float 6s ease-in-out infinite",
        shimmer: "shimmer 2.5s linear infinite",
      },
    },
  },
  plugins: [],
};
