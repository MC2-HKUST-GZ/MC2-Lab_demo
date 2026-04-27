import fs from 'fs/promises';
import path from 'path';
import matter from 'gray-matter';
import { glob } from 'glob';
import type { PublicationRecord, PublicationSubmission } from '../src/types/common';

const ROOT = process.cwd();
const SUBMISSION_DIR = path.join(ROOT, 'content/submissions/publications');
const IMAGE_DIR = path.join(SUBMISSION_DIR, 'images');
const OUTPUT_FILE = path.join(ROOT, 'src/data/publications.generated.json');

const requiredFields: Array<keyof PublicationSubmission> = [
  'title',
  'conference',
  'photo',
  'representative_work',
  'link'
];

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}

function assertValidUrl(url: string, filePath: string): void {
  try {
    const parsed = new URL(url);
    if (!['http:', 'https:'].includes(parsed.protocol)) {
      throw new Error('URL must use http/https');
    }
  } catch (error) {
    throw new Error(`[${filePath}] invalid link: ${url}. ${(error as Error).message}`);
  }
}

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function toRecord(filePath: string, data: PublicationSubmission): PublicationRecord {
  const id = path.basename(filePath, path.extname(filePath));
  return {
    id,
    title: data.title.trim(),
    conference: data.conference.trim(),
    photo: data.photo.trim(),
    representativeWork: data.representative_work.trim(),
    link: data.link.trim()
  };
}

async function run(): Promise<void> {
  const markdownFiles = await glob('*.md', {
    cwd: SUBMISSION_DIR,
    absolute: true,
    ignore: ['README.md']
  });

  const seenLinks = new Set<string>();
  const output: PublicationRecord[] = [];

  for (const mdFile of markdownFiles.sort()) {
    const source = await fs.readFile(mdFile, 'utf8');
    const { data } = matter(source);

    for (const field of requiredFields) {
      if (!isNonEmptyString(data[field])) {
        throw new Error(`[${path.relative(ROOT, mdFile)}] missing required field: ${field}`);
      }
    }

    const submission = data as PublicationSubmission;
    assertValidUrl(submission.link, path.relative(ROOT, mdFile));

    const imagePath = path.join(IMAGE_DIR, submission.photo);
    if (!(await fileExists(imagePath))) {
      throw new Error(
        `[${path.relative(ROOT, mdFile)}] photo not found: ${submission.photo} (expected in content/submissions/publications/images/)`
      );
    }

    if (seenLinks.has(submission.link.trim())) {
      throw new Error(`[${path.relative(ROOT, mdFile)}] duplicate link: ${submission.link}`);
    }
    seenLinks.add(submission.link.trim());

    output.push(toRecord(mdFile, submission));
  }

  await fs.mkdir(path.dirname(OUTPUT_FILE), { recursive: true });
  await fs.writeFile(OUTPUT_FILE, `${JSON.stringify(output, null, 2)}\n`, 'utf8');

  console.log(`Generated ${output.length} publication record(s): ${path.relative(ROOT, OUTPUT_FILE)}`);
}

run().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
