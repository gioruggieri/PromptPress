import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  outputFileTracingIncludes: {
    "/api/pdf": ["./node_modules/@sparticuz/chromium/**/*"],
    "/api/pdf/route": ["./node_modules/@sparticuz/chromium/**/*"],
    "/api/docx-pandoc": ["./.pandoc/**/*"],
    "/api/docx-pandoc/route": ["./.pandoc/**/*"],
    "/app/api/docx-pandoc/route": ["./.pandoc/**/*"],
  },
};

export default nextConfig;
