import * as docx from 'docx';

export async function buildDocx(
  code: string,
  assessment: object,
  questionnaire: object,
  commissionBanner: string,
  unitLogo: string
): Promise<Blob> {
  // we expect code that contains a function: function generateDocument(): docx.Document { ... }
  const injectedCode = `
    const ASSESSMENT = ${JSON.stringify(assessment)};
    const QUESTIONNAIRE = ${JSON.stringify(questionnaire)};
    const commissionBannerBase64 = \`${commissionBanner}\`;
    const unitLogoBase64 = \`${unitLogo}\`;
  `;
  // Prepend the injected variables to the user's original code.
  const fullCode = `${injectedCode}\n${code}`;

  const trimmedCode = fullCode.trim();
  const codeWithoutImportAndExportStatements = trimmedCode
    .split('\n')
    .filter((line) => !line.startsWith('import '))
    .filter((line) => !line.startsWith('export '))
    .join('\n');
  const codeStrictMode = `"use strict";\n${codeWithoutImportAndExportStatements}`;
  const codeWithReturnFunctionInvocation = `${codeStrictMode}\nreturn generateDocument();\n`;
  const userFunction = new Function('docx', codeWithReturnFunctionInvocation);
  const doc = userFunction(docx);
  if (!(doc instanceof docx.Document)) {
    throw new Error(
      'The function generateDocument should return a docx.Document instance'
    );
  }
  return await docx.Packer.toBlob(doc);
}
