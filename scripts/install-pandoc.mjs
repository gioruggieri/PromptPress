import { mkdir, rm, stat } from "node:fs/promises";
import { createWriteStream } from "node:fs";
import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";
import https from "node:https";

const VERSION = process.env.PANDOC_VERSION || "3.1.11.1";
const ROOT = process.cwd();
const OUT_DIR = path.join(ROOT, ".pandoc");

const platform = process.platform;
const arch = process.arch;
const skipInstall =
  process.env.PANDOC_SKIP_INSTALL === "1" ||
  (platform === "win32" && process.env.PANDOC_FORCE_INSTALL !== "1");

const resolveDownload = () => {
  if (platform === "linux" && arch === "x64") {
    return {
      filename: `pandoc-${VERSION}-linux-amd64.tar.gz`,
      innerDir: `pandoc-${VERSION}`,
      binary: "pandoc",
      binSubdir: "bin",
    };
  }
  if (platform === "darwin" && arch === "x64") {
    return {
      filename: `pandoc-${VERSION}-macOS.zip`,
      innerDir: `pandoc-${VERSION}`,
      binary: "pandoc",
      binSubdir: "bin",
    };
  }
  if (platform === "darwin" && arch === "arm64") {
    return {
      filename: `pandoc-${VERSION}-macOS-arm64.zip`,
      innerDir: `pandoc-${VERSION}`,
      binary: "pandoc",
      binSubdir: "bin",
    };
  }
  if (platform === "win32") {
    return {
      filename: `pandoc-${VERSION}-windows-x86_64.zip`,
      innerDir: `pandoc-${VERSION}`,
      binary: "pandoc.exe",
      binSubdir: "bin",
    };
  }
  return null;
};

const run = (cmd, args, options = {}) =>
  new Promise((resolve, reject) => {
    const proc = spawn(cmd, args, { stdio: "inherit", ...options });
    proc.on("error", reject);
    proc.on("close", (code) =>
      code === 0 ? resolve() : reject(new Error(`${cmd} exited with ${code}`)),
    );
  });

const download = (url, dest, redirects = 0) =>
  new Promise((resolve, reject) => {
    const stream = createWriteStream(dest);
    https
      .get(url, (res) => {
        const status = res.statusCode ?? 0;
        const location = res.headers.location;
        if (status >= 300 && status < 400 && location) {
          if (redirects > 4) {
            reject(new Error("Too many redirects"));
            return;
          }
          stream.close(() => {
            download(location, dest, redirects + 1).then(resolve).catch(reject);
          });
          return;
        }
        if (status >= 400) {
          reject(new Error(`Download failed: ${status}`));
          return;
        }
        res.pipe(stream);
        stream.on("finish", () => stream.close(resolve));
      })
      .on("error", reject);
  });

const ensureDir = async (dir) => {
  try {
    await mkdir(dir, { recursive: true });
  } catch {
    // ignore
  }
};

const main = async () => {
  if (skipInstall) {
    console.log("Pandoc install skipped.");
    return;
  }
  const target = resolveDownload();
  if (!target) {
    console.warn("Pandoc install skipped: unsupported platform", platform, arch);
    return;
  }

  const binPath = path.join(OUT_DIR, target.binary);
  if (existsSync(binPath)) return;

  await ensureDir(OUT_DIR);

  const url = `https://github.com/jgm/pandoc/releases/download/${VERSION}/${target.filename}`;
  const tmp = path.join(os.tmpdir(), target.filename);

  await download(url, tmp);

  if (target.filename.endsWith(".tar.gz")) {
    await run("tar", ["-xzf", tmp, "-C", OUT_DIR]);
  } else if (target.filename.endsWith(".zip")) {
    if (platform === "win32") {
      await run("powershell", [
        "-NoProfile",
        "-Command",
        `Expand-Archive -LiteralPath "${tmp}" -DestinationPath "${OUT_DIR}" -Force`,
      ]);
    } else {
      await run("unzip", ["-o", tmp, "-d", OUT_DIR]);
    }
  }

  const candidates = [
    path.join(OUT_DIR, target.innerDir, target.binary),
    path.join(OUT_DIR, target.innerDir, target.binSubdir, target.binary),
  ];
  const extracted = candidates.find((candidate) => existsSync(candidate));
  if (!extracted) {
    throw new Error("Pandoc binary not found after extraction");
  }

  await rm(binPath, { force: true });
  if (platform === "win32") {
    await run("cmd", ["/c", "copy", "/y", extracted, binPath]);
  } else {
    await run("ln", ["-s", extracted, binPath]);
    await run("chmod", ["+x", extracted]);
  }

  try {
    await stat(binPath);
  } catch {
    throw new Error("Pandoc install failed");
  }
};

main().catch((err) => {
  console.error("Pandoc install failed:", err.message);
  process.exit(1);
});
