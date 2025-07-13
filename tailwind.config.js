/** @type {import('tailwindcss').Config} */
module.exports = {
    darkMode: ['class'],
    content: [
    './src/pages/**/*.{js,ts,jsx,tsx,mdx}',
    './src/components/**/*.{js,ts,jsx,tsx,mdx}',
    './src/app/**/*.{js,ts,jsx,tsx,mdx}',
  ],
  theme: {
  	extend: {
  		colors: {
  			background: 'hsl(var(--background))',
  			foreground: 'hsl(var(--foreground))',
  			card: {
  				DEFAULT: 'hsl(var(--card))',
  				foreground: 'hsl(var(--card-foreground))'
  			},
  			popover: {
  				DEFAULT: 'hsl(var(--popover))',
  				foreground: 'hsl(var(--popover-foreground))'
  			},
  			primary: {
  				DEFAULT: 'hsl(var(--primary))',
  				foreground: 'hsl(var(--primary-foreground))'
  			},
  			secondary: {
  				DEFAULT: 'hsl(var(--secondary))',
  				foreground: 'hsl(var(--secondary-foreground))'
  			},
  			muted: {
  				DEFAULT: 'hsl(var(--muted))',
  				foreground: 'hsl(var(--muted-foreground))'
  			},
  			accent: {
  				DEFAULT: 'hsl(var(--accent))',
  				foreground: 'hsl(var(--accent-foreground))'
  			},
  			destructive: {
  				DEFAULT: 'hsl(var(--destructive))',
  				foreground: 'hsl(var(--destructive-foreground))'
  			},
  			border: 'hsl(var(--border))',
  			input: 'hsl(var(--input))',
  			ring: 'hsl(var(--ring))',
  			chart: {
  				'1': 'hsl(var(--chart-1))',
  				'2': 'hsl(var(--chart-2))',
  				'3': 'hsl(var(--chart-3))',
  				'4': 'hsl(var(--chart-4))',
  				'5': 'hsl(var(--chart-5))'
  			},
			// Gray palette
			gray: {
				"010": "#F2F5F7",
				"020": "#DCE3E8",
				"030": "#C1CCD6",
				"040": "#9FB1BD",
				"050": "#7A909E",
				"060": "#5B7282",
				"070": "#3E5463",
				"080": "#2A3F4D",
				"090": "#1C2B36",
				"100": "#0E171F",
				"110": "#090F14",
			  },
			  // Blue palette
			  blue: {
				"010": "#F0F4FA",
				"020": "#D4E4FA",
				"030": "#ADCCF7",
				"040": "#75B1FF",
				"050": "#3D8DF5",
				"060": "#186ADE",
				"070": "#0D4EA6",
				"080": "#103A75",
				"090": "#11294D",
				"100": "#0D1826",
			  },
			  // Aqua palette
			  aqua: {
				"010": "#EBF3F7",
				"020": "#C9E7F5",
				"030": "#8BD3F7",
				"040": "#48B8F0",
				"050": "#1195D6",
				"060": "#0073BA",
				"070": "#08548A",
				"080": "#0E3D66",
				"090": "#0C2B45",
				"100": "#0B1724",
			  },
			  // Green palette
			  green: {
				"010": "#EBF7ED",
				"020": "#C7EBD1",
				"030": "#88DBA8",
				"040": "#43C478",
				"050": "#16A163",
				"060": "#077D55",
				"070": "#075E45",
				"080": "#094536",
				"090": "#092E25",
				"100": "#081A15",
			  },
			  // Red palette
			  red: {
				"010": "#FCF3F2",
				"020": "#FACDD9",
				"030": "#FABBB4",
				"040": "#FC9086",
				"050": "#FA5343",
				"060": "#D91F11",
				"070": "#A1160A",
				"080": "#75160C",
				"090": "#4F150F",
				"100": "#24120C",
			  },
			  // Yellow palette
			  yellow: {
				"010": "#FAF6CF",
				"020": "#F7E379",
				"030": "#F5C518",
				"040": "#D9A514",
				"050": "#B3870E",
				"060": "#946613",
				"070": "#70491C",
				"080": "#54341F",
				"090": "#38251B",
				"100": "#1C1613",
			  },
			  // Purple palette
			  purple: {
				"010": "#F7F2FC",
				"020": "#EADCFC",
				"030": "#DABEFA",
				"040": "#C89AFC",
				"050": "#AC71F0",
				"060": "#8F49DE",
				"070": "#6B30AB",
				"080": "#4C277D",
				"090": "#331F4D",
				"100": "#1C1229",
			  },
			  // Teal palette
			  teal: {
				"010": "#EBF5F4",
				"020": "#BEEBE7",
				"030": "#86D9D4",
				"040": "#4EBFB9",
				"050": "#279C9C",
				"060": "#167B7D",
				"070": "#155C5E",
				"080": "#124241",
				"090": "#102E2D",
				"100": "#0C1A19",
			  },
  		},
  		borderRadius: {
  			lg: 'var(--radius)',
  			md: 'calc(var(--radius) - 2px)',
  			sm: 'calc(var(--radius) - 4px)'
  		}
  	}
  },
  plugins: [require("tailwindcss-animate")],
}; 