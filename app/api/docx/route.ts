import { NextRequest, NextResponse } from 'next/server';
import fs from 'node:fs/promises';
import path from 'node:path';
import { v4 as uuidv4 } from 'uuid';
import { buildDocx } from '@/lib/build-docx';
import { TMP_DIR } from '@/lib/file-system';

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const code = formData.get('code') as string | null;
    if (!code) {
      return NextResponse.json({ error: 'Code is required' }, { status: 400 });
    }
    let docxBlob: Blob;

    // 1. Get the full path to the JSON files
    const assessmentPath = path.join(
      process.cwd(),
      'public',
      'assets',
      'Complete_VAPP',
      'assessment.json'
    );
    const questionnairePath = path.join(
      process.cwd(),
      'public',
      'assets',
      'questionnaire.json'
    );
    // 2. Read the raw text content using the basic 'fs' module function
    const assessmentJson = await fs.readFile(assessmentPath, 'utf-8');
    const questionnaireJson = await fs.readFile(questionnairePath, 'utf-8');
    // 3. Parse the JSON strings into objects
    const assessment = JSON.parse(assessmentJson);
    const questionnaire = JSON.parse(questionnaireJson);
    // 4. Also add the logos
    const commissionBannerPath = path.join(
      process.cwd(),
      'public',
      'assets',
      'commission-banner.b64'
    );
    const unitLogoPath = path.join(
      process.cwd(),
      'public',
      'assets',
      'unit-logo.b64'
    );
    const commissionBanner = await fs.readFile(commissionBannerPath, 'utf-8');
    const unitLogo = await fs.readFile(unitLogoPath, 'utf-8');

    try {
      docxBlob = await buildDocx(
        code,
        assessment,
        questionnaire,
        commissionBanner,
        unitLogo
      );
    } catch (err) {
      console.error(err);
      return NextResponse.json(
        { error: `Error generating docx: ${err}` },
        { status: 400 }
      );
    }
    if (!docxBlob) {
      return NextResponse.json(
        { error: 'Failed to generate docx' },
        { status: 400 }
      );
    }
    const fileId = uuidv4(); // randomly generated file ID
    await fs.mkdir(TMP_DIR, { recursive: true });
    const filePath = path.join(TMP_DIR, `${fileId}.docx`);
    const fileBuffer = await docxBlob.arrayBuffer();
    await fs.writeFile(filePath, Buffer.from(fileBuffer));
    return NextResponse.json({ fileId }, { status: 200 });
  } catch (err) {
    console.error(err);
    return NextResponse.json(
      { error: `Internal server error` },
      { status: 500 }
    );
  }
}
