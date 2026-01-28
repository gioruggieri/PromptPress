import { NextResponse } from "next/server";
import { spawn } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";

export const runtime = "nodejs";

const resolvePandocPath = () => {
  const envPath = process.env.PANDOC_PATH;
  if (envPath && envPath.trim()) return envPath.trim();
  const localBinary = path.join(
    process.cwd(),
    ".pandoc",
    process.platform === "win32" ? "pandoc.exe" : "pandoc",
  );
  if (existsSync(localBinary)) return localBinary;
  const installedBinary = path.join(
    process.cwd(),
    ".pandoc",
    `pandoc-${process.env.PANDOC_VERSION || "3.1.11.1"}`,
    "bin",
    process.platform === "win32" ? "pandoc.exe" : "pandoc",
  );
  if (existsSync(installedBinary)) return installedBinary;
  return "pandoc";
};

const runPandoc = async (markdown: string) => {
  const tmpDir = await mkdtemp(path.join(os.tmpdir(), "promptpress-"));
  const inputPath = path.join(tmpDir, "input.md");
  const outputPath = path.join(tmpDir, "output.docx");

  try {
    await writeFile(inputPath, markdown, "utf8");

    const pandocPath = resolvePandocPath();
    const args = [
      "-f",
      "gfm+tex_math_dollars+pipe_tables+task_lists+strikeout+autolink_bare_uris",
      "-t",
      "docx",
      "-o",
      outputPath,
      inputPath,
    ];

    const { code, stderr } = await new Promise<{ code: number; stderr: string }>(
      (resolve, reject) => {
        const proc = spawn(pandocPath, args, { windowsHide: true });
        let err = "";

        proc.stderr.on("data", (chunk) => {
          err += chunk.toString();
        });
        proc.on("error", (error) => reject(error));
        proc.on("close", (exitCode) => resolve({ code: exitCode ?? 1, stderr: err }));
      },
    );

    if (code !== 0) {
      const message = stderr.trim() || `Pandoc exited with code ${code}`;
      throw new Error(message);
    }

    return await readFile(outputPath);
  } finally {
    await rm(tmpDir, { recursive: true, force: true });
  }
};

export async function POST(request: Request) {
  try {
    const { markdown } = await request.json();

    if (!markdown || typeof markdown !== "string") {
      return NextResponse.json({ error: "Missing markdown" }, { status: 400 });
    }

    const fileBuffer = await runPandoc(markdown);
    const mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const body = new Blob([fileBuffer], { type: mime });

    return new NextResponse(body, {
      status: 200,
      headers: {
        "Content-Type": mime,
        "Content-Disposition": 'attachment; filename="promptpress-export-pandoc.docx"',
      },
    });
  } catch (error) {
    let message = error instanceof Error ? error.message : "Pandoc error";
    const code = (error as NodeJS.ErrnoException | undefined)?.code;
    if (code === "ENOENT") {
      message =
        "Pandoc not found. Set PANDOC_PATH or provide a local .pandoc/pandoc binary.";
    }
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
