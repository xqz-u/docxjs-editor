import * as docx from 'docx';

const flattenQuestionnaire = (modularQuestionnaire) => {
  const flatQuestionnaire = {};
  Object.values(modularQuestionnaire).forEach((module) => {
    Object.assign(flatQuestionnaire, module.questions);
  });
  return flatQuestionnaire;
};

const collectUniqueRiskDimensions = (questionnaire) => {
  const uniqueDimensions = new Set();

  Object.values(questionnaire).forEach((question) => {
    question.answers.forEach((answer) => {
      answer.risk_dimensions.forEach((dimension) => {
        uniqueDimensions.add(dimension);
      });
    });
  });

  return Array.from(uniqueDimensions);
};

const findQuestionIdBySerial = (querySerial, questions) =>
  Number(
    Object.entries(questions).find(
      ([, { serial }]) => serial === querySerial
    )[0]
  );

const computeScoreVector = (userAnswers, questionnaire, riskDimensions) => {
  // console.log('User Answers');
  // console.log(userAnswers);

  const scoreVector = Object.fromEntries(riskDimensions.map((dim) => [dim, 0]));
  const countsVector = { ...scoreVector };

  Object.entries(userAnswers).forEach(([qid, answers]) => {
    // Free text answers do not have a value
    // assert question.type === "FREE_TEXT";
    if (!Array.isArray(answers)) return;

    const question = questionnaire[qid];
    const questionValue = question.value ?? 1;
    // console.log(
    //   `Question ${question.serial} (${qid}) :: value ${questionValue}`,
    // );

    answers.forEach((answerId) => {
      const answer = question.answers.filter(({ id }) => id === answerId)[0];
      const { value } = answer;
      // console.log(`Answer ${answerId}: value ${value} dims:`);
      // console.log(answer.risk_dimensions);

      // Above, we filtered out answer values that are null, 'NA' or 'NK'
      if (typeof value !== 'number') return;

      answer.risk_dimensions.forEach((dim) => {
        scoreVector[dim] += value / questionValue;
        countsVector[dim] += 1;
      });
    });
  });

  // prettyPrintObject(countsVector);
  // prettyPrintObject(scoreVector);

  const weightedScoreVector = Object.fromEntries(
    Object.entries(scoreVector).map(([dim, score]) => [
      dim,
      // handle case of division by 0
      countsVector[dim] > 0 ? score / countsVector[dim] : 0,
    ])
  );
  return [weightedScoreVector, countsVector, scoreVector];
};

/**
 * @returns {docx.Document}
 * @see https://docx.js.org/
 */
function generateDocument() {
  const showFreeTextAnswer = (querySerial, assessment, questionnaire) =>
    assessment.content.questionnaireProgress.answers[
      findQuestionIdBySerial(querySerial, questionnaire)
    ];

  const showAllAnswersToChoiceQuestion = (
    serial,
    assessment,
    questionnaire
  ) => {
    const questionId = findQuestionIdBySerial(serial, questionnaire);
    const possibleAnswers = questionnaire[questionId].answers;
    const selectedAnswerIds =
      assessment.content.questionnaireProgress.answers[questionId];
    return possibleAnswers.map(({ id, content }, index) => {
      const isSelectedAnswer = selectedAnswerIds.includes(id);
      return new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: content,
            underline: isSelectedAnswer,
            bold: isSelectedAnswer,
          }),
        ],
        bullet: { level: 0 },
        spacing: { after: index === possibleAnswers.length - 1 ? 0 : 50 },
      });
    });
  };

  const showAnswerToChoiceQuestion = (serial, assessment, questionnaire) => {
    const questionId = findQuestionIdBySerial(serial, questionnaire);
    const [answerId] =
      assessment.content.questionnaireProgress.answers[questionId];
    const answer = questionnaire[questionId].answers.find(
      ({ id }) => id === answerId
    );
    if (answer !== undefined)
      return new docx.Paragraph({ text: answer.content });
  };

  const createScoreTable = (scores) => {
    const rows = Object.entries(scores).map(
      ([dim, score]) =>
        new docx.TableRow({
          children: [
            new docx.TableCell({
              children: [new docx.Paragraph(dim)],
              width: { size: 50, type: docx.WidthType.PERCENTAGE },
            }),
            new docx.TableCell({
              children: [new docx.Paragraph(score.toFixed(4))],
              width: { size: 50, type: docx.WidthType.PERCENTAGE },
            }),
          ],
        })
    );

    return new docx.Table({
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: 'Risk Dimension', bold: true }),
                  ],
                }),
              ],
            }),
            new docx.TableCell({
              children: [
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({ text: 'Weighted Score', bold: true }),
                  ],
                }),
              ],
            }),
          ],
        }),
        ...rows,
      ],
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
    });
  };

  const createTitlePage = (assessment, questionnaire) => {
    return [
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Vulnerability Assessment for',
            bold: true,
            size: 44, // 22pt font size
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 1500 }, // Spacing from top of the page
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: showFreeTextAnswer('3.1', assessment, questionnaire),
            size: 44,
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 800 }, // Spacing between lines
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'located at',
            italics: true,
            size: 36, // 18pt font size
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 800 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: showFreeTextAnswer('2.1.1', assessment, questionnaire),
            size: 44,
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 800 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'on date(s)',
            italics: true,
            size: 36,
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 800 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: showFreeTextAnswer('3.6.1', assessment, questionnaire),
            size: 44,
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 800 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: `Assessment date: ${new Date().toISOString()}`,
            size: 28, // 14pt font size
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 1500 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'This report was automatically generated by the Vulnerability Assessment application V1.0 provided by DG HOME Counter-Terrorism Unit.',
            size: 28,
          }),
        ],
        alignment: docx.AlignmentType.CENTER,
        spacing: { before: 400 },
      }),
    ];
  };

  const createDisclaimers = () => {
    // Content for the first box
    const firstBoxParagraphs = [
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Acts covered by this application',
            bold: true,
          }),
        ],
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'This application focuses and covers terrorist acts and other antagonistic behaviour, intentionally malicious and illegal.',
            italics: true,
          }),
        ],
        spacing: { after: 100 },
      }),
      new docx.Paragraph({
        text: 'They are actor-driven, i.e. they are dependent on someone actively trying to avoid security measures to achieve objectives. The determining factor is the perpetrator’s intention to commit an act. Such acts are committed with the intention of striking fear into a population or a group within a population; forcing public authorities to take or avoid taking previously decided measures; destabilising or destroying basic political, constitutional, economic, or social structures in a state or between states.',
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'This application does not cover accidents or near-accidents (unintentional occurrences that happen suddenly and lead to some form of injury or damage).',
            italics: true,
          }),
        ],
        spacing: { after: 100 },
      }),
      new docx.Paragraph({
        text: 'This refers e.g. to crowd surges and consequences of stumbling/falls, road traffic accidents, extreme weather phenomena, collapsing structures, medical emergencies due to heat, dehydration, stroke, etc., consequences of alcohol/drug abuse, food poisoning and infectious disease spread in crowd. Such events are not subject of analysis within the context of this application.',
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'The application does not also cover common crime-related acts other than terrorist acts.',
            italics: true,
          }),
        ],
        spacing: { after: 100 },
      }),
      new docx.Paragraph({
        text: 'It is not uncommon for crimes to be committed in connection with events, since many people congregate in a small space. Visitors may also become intoxicated, which influences both their propensity to commit crimes and the risk of their becoming victims of crime. This refers e.g. to theft, pickpocketing, sexual aggressions, assault, and battery. Such events are not subject of analysis within the context of this application.',
      }),
    ];

    // Content for the second box
    const secondBoxParagraphs = [
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Legal context and responsibilities of private and public actors in the management of safety and security of public events and public spaces',
            bold: true,
          }),
        ],
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        text: 'In Europe, roles and responsibilities of the public and private actors, involved in the management of safety and security of public events and public spaces, vary depending on national and sometimes even regional legislation.',
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        text: "During the evaluation and assessment process, the relevant local legal context and the responsibilities should be carefully verified, considered, and complied with to the extent possible. The application's questions, advice and guidance in this context are generic and cannot consider the specific local legal context. This should be ensured by the assessment team's comments and recommendations.",
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        text: 'The application aims to foster exchange and dialogue between the various actors, regardless of their specific skills and roles in managing local safety and security, so that all the main aspects of the process are considered. It focuses on the prevention and management of criminal acts, particularly the event of a terrorist attack. Safety and security may overlap, for example, if during a terrorist attack at the event venue, law enforcement authorities must intervene, and the event organizer manages the evacuation of the public through its emergency management staff.',
      }),
    ];

    // Define the border style for the cells
    const cellBorders = {
      top: { style: docx.BorderStyle.SINGLE, size: 6, color: 'auto' },
      bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: 'auto' },
      left: { style: docx.BorderStyle.SINGLE, size: 6, color: 'auto' },
      right: { style: docx.BorderStyle.SINGLE, size: 6, color: 'auto' },
    };

    // Create the first table (box)
    const table1 = new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              children: firstBoxParagraphs,
              borders: cellBorders,
              margins: { top: 200, bottom: 200, left: 200, right: 200 },
            }),
          ],
        }),
      ],
    });

    // Create the second table (box)
    const table2 = new docx.Table({
      width: { size: 100, type: docx.WidthType.PERCENTAGE },
      rows: [
        new docx.TableRow({
          children: [
            new docx.TableCell({
              children: secondBoxParagraphs,
              borders: cellBorders,
              margins: { top: 200, bottom: 200, left: 200, right: 200 },
            }),
          ],
        }),
      ],
    });

    return [
      new docx.Paragraph({ text: '' }), // Spacer paragraph between boxes
      table1,
      new docx.Paragraph({ text: '' }), // Spacer paragraph between boxes
      table2,
    ];
  };

  const createVenueInformation = (assessment, questionnaire) => {
    return [
      new docx.Paragraph({
        text: 'Venue Information',
        heading: docx.HeadingLevel.HEADING_1,
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Description of the venue location',
            bold: true,
          }),
        ],
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        text: showFreeTextAnswer('2.1.1', assessment, questionnaire),
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Sectors applicable for this location',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      ...showAllAnswersToChoiceQuestion('2.2', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Venue location type',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      showAnswerToChoiceQuestion('2.7', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Venue type characteristics',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      new docx.Paragraph({
        text: showFreeTextAnswer('2.7.1', assessment, questionnaire),
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Overall size (gross area)',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      showAnswerToChoiceQuestion('2.13', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Characteristics of surrounding area',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      showAnswerToChoiceQuestion('2.10', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Venue maps (???????)',
            bold: true,
          }),
        ],
        spacing: { before: 200 },
      }),
    ];
  };

  const createEventInformation = (assessment, questionnaire) => {
    return [
      new docx.Paragraph({ text: '' }),
      new docx.Paragraph({
        text: 'Event Information',
        heading: docx.HeadingLevel.HEADING_1,
        spacing: { before: 200, after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Main activities (program/agenda)',
            bold: true,
          }),
        ],
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        text: showFreeTextAnswer('3.2', assessment, questionnaire),
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Frequency',
            bold: true,
          }),
        ],
        spacing: { after: 200 },
      }),
      showAnswerToChoiceQuestion('3.5', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Duration',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      showAnswerToChoiceQuestion('3.6', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Event opening and closing times for the public (day/time)',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      new docx.Paragraph({
        text: showFreeTextAnswer('3.8', assessment, questionnaire),
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Maximum expected attendance',
            bold: true,
          }),
        ],
        spacing: { after: 200 },
      }),
      showAnswerToChoiceQuestion('3.9', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Expected average crowd density (participants/m²)',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      showAnswerToChoiceQuestion('3.11', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'At what level is the event known?',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      showAnswerToChoiceQuestion('3.16', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "Activity/event's character",
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      showAnswerToChoiceQuestion('3.17', assessment, questionnaire),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Media coverage and live broadcasting',
            bold: true,
          }),
        ],
        spacing: { before: 200, after: 200 },
      }),
      showAnswerToChoiceQuestion('3.19', assessment, questionnaire),
    ];
  };

  const newPage = () =>
    new docx.Paragraph({ children: [new docx.PageBreak()] });

  // Fetch the JSON data
  // const assessmentResponse = await fetch('/assets/Complete_VAPP/assessment.json');
  // const ASSESSMENT = await assessmentResponse.json();

  // const questionnaireResponse = await fetch('/assets/questionnaire.json');
  // const QUESTIONNAIRE = await questionnaireResponse.json();

  const FLAT_QUESTIONNAIRE = flattenQuestionnaire(QUESTIONNAIRE);
  const UNIQUE_RISK_DIMENSIONS =
    collectUniqueRiskDimensions(FLAT_QUESTIONNAIRE);

  // Your code goes here
  // 1. COMPUTE SCORES
  // We get the answers from the loaded assessment file
  const userAnswers = ASSESSMENT.content.questionnaireProgress.answers;
  // Note: computeScoreVector returns an array. We want the first element.
  const [weightedScoreVector] = computeScoreVector(
    userAnswers,
    FLAT_QUESTIONNAIRE,
    UNIQUE_RISK_DIMENSIONS
  );

  // 2. BUILD THE DOCUMENT STRUCTURE
  const doc = new docx.Document({
    sections: [
      {
        features: { updateFields: true },
        headers: {
          default: new docx.Header({
            children: [
              // Use a SINGLE table for the layout
              new docx.Table({
                width: { size: '100%', type: docx.WidthType.PERCENTAGE },
                // Apply the blue border directly to the main table
                borders: {
                  top: {
                    style: docx.BorderStyle.SINGLE,
                    size: 24,
                    color: '004494',
                  },
                  bottom: {
                    style: docx.BorderStyle.SINGLE,
                    size: 24,
                    color: '004494',
                  },
                  left: {
                    style: docx.BorderStyle.SINGLE,
                    size: 24,
                    color: '004494',
                  },
                  right: {
                    style: docx.BorderStyle.SINGLE,
                    size: 24,
                    color: '004494',
                  },
                },
                rows: [
                  new docx.TableRow({
                    children: [
                      // --- Column 1: Left Logo ---
                      new docx.TableCell({
                        width: { size: 25, type: docx.WidthType.PERCENTAGE },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.ImageRun({
                                data: commissionBannerBase64.replace(
                                  'data:image/png;base64,',
                                  ''
                                ),
                                type: 'png',
                                transformation: { width: 150, height: 100 },
                              }),
                            ],
                          }),
                        ],
                        verticalAlign: docx.VerticalAlign.CENTER,
                      }),
                      // --- Column 2: Centered Text ---
                      new docx.TableCell({
                        width: { size: 50, type: docx.WidthType.PERCENTAGE },
                        children: [
                          new docx.Paragraph({
                            text: 'Vulnerability Assessment for Public Spaces',
                            alignment: docx.AlignmentType.CENTER,
                          }),
                        ],
                        verticalAlign: docx.VerticalAlign.CENTER,
                      }),
                      // --- Column 3: Right Logo ---
                      new docx.TableCell({
                        width: { size: 25, type: docx.WidthType.PERCENTAGE },
                        children: [
                          new docx.Paragraph({
                            children: [
                              new docx.ImageRun({
                                data: unitLogoBase64.replace(
                                  'data:image/png;base64,',
                                  ''
                                ),
                                type: 'png',
                                transformation: { width: 100, height: 100 },
                              }),
                            ],
                            alignment: docx.AlignmentType.CENTER,
                          }),
                        ],
                        verticalAlign: docx.VerticalAlign.CENTER,
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        },
        children: [
          ...createTitlePage(ASSESSMENT, FLAT_QUESTIONNAIRE),
          newPage(),
          ...createDisclaimers(ASSESSMENT),
          newPage(),
          new docx.TableOfContents('Table of Contents', {
            hyperlink: true,
            headingStyleRange: '1-5',
          }),
          ...createVenueInformation(ASSESSMENT, FLAT_QUESTIONNAIRE),
          newPage(),
          ...createEventInformation(ASSESSMENT, FLAT_QUESTIONNAIRE),
          newPage(),
          // --- Score Summary Section ---
          new docx.Paragraph({
            text: 'Risk Score Summary',
            heading: docx.HeadingLevel.HEADING_1,
            pageBreakBefore: true, // Start this section on a new page
            spacing: { before: 200, after: 200 },
          }),
          new docx.Paragraph({
            text: 'The following chart provides a visual representation of the overall risk score.',
          }),
          new docx.Paragraph({
            children: [
              //   new docx.ImageRun({
              //     // We need to remove the data URI prefix for the docx library
              //     data: chartImageBase64.replace("data:image/png;base64,", ""),
              //     transformation: {
              //       width: 500,
              //       height: 100,
              //     },
              //     type: "png",
              //   }),
            ],
          }),
          new docx.Paragraph({ text: '' }), // Spacer
          new docx.Paragraph({
            text: 'The following table presents the weighted scores calculated for each risk dimension based on the provided answers.',
          }),
          new docx.Paragraph({ text: '' }), // Spacer
          createScoreTable(weightedScoreVector),
        ],
      },
    ],
  });

  return doc;
}

export default generateDocument;
