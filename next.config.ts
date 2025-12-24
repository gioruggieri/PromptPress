import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  outputFileTracingIncludes: {
    "/api/pdf": ["./node_modules/@sparticuz/chromium/**/*"],
  },
};

export default nextConfig;
