/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
  ],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        brand: {
          50: '#f0fdfa',
          100: '#ccfbf1',
          200: '#99f6e4',
          300: '#5eead4',
          400: '#2dd4bf', // Accento vibrante (Dark mode)
          500: '#14b8a6',
          600: '#0d9488', // Accento sólido (Light mode)
          700: '#0f766e',
          800: '#115e59',
          900: '#134e4a',
          950: '#042f2e',
        },
        primary: {
          DEFAULT: '#0d9488', // Cambiamos a un tono más profundo por defecto
          dark: '#2dd4bf',
        },
        background: {
          light: '#f8fafc',
          dark: '#0f172a',
        },
        surface: {
          light: '#ffffff',
          dark: '#1e293b',
        }
      },
      fontFamily: {
        display: ['"Spline Sans"', 'system-ui', 'sans-serif'],
      },
      boxShadow: {
        'soft': '0 2px 15px -3px rgba(0, 0, 0, 0.07), 0 4px 6px -2px rgba(0, 0, 0, 0.05)',
        'brand': '0 10px 15px -3px rgba(13, 148, 136, 0.2)',
      }
    },
  },
  plugins: [],
}
