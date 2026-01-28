import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  outputFileTracingIncludes: {
    "/api/pdf": ["./node_modules/@sparticuz/chromium/**/*"],
    "/api/pdf/route": ["./node_modules/@sparticuz/chromium/**/*"],
    "/api/docx-pandoc": [
      "./node_modules/pandoc-bin/**/*",
      "./node_modules/bin-wrapper/**/*",
      "./node_modules/os-filter-obj/**/*",
    ],
    "/api/docx-pandoc/route": [
      "./node_modules/pandoc-bin/**/*",
      "./node_modules/bin-wrapper/**/*",
      "./node_modules/os-filter-obj/**/*",
    ],
  },
};

export default nextConfig;
