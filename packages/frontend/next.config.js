/** @type {import('next').NextConfig} */
const nextConfig = {
  async rewrites() {
    // In development, proxy API calls to the FastAPI server
    // In production, Vercel handles this automatically
    if (process.env.NODE_ENV === 'development') {
    return [
      {
        source: '/api/:path*',
        destination: 'http://localhost:8000/api/:path*',
      },
    ];
    }
    return [];
  },
};

module.exports = nextConfig; 