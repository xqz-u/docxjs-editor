import * as docx from 'docx';

const evaluateDefaultCheck = (
  { question_id, op },
  storedAnswers,
  questions
) => {
  if (!(question_id in storedAnswers)) return false;

  const { default_answer } = questions[question_id];

  let normalizedAnswer;
  let normalizedDefault;
  const currentAnswer = storedAnswers[question_id];

  if (Array.isArray(currentAnswer) && Array.isArray(default_answer)) {
    // Avoid difference in comparisons due to answer ordering
    normalizedAnswer = _.sortBy(currentAnswer);
    normalizedDefault = _.sortBy(default_answer);
  } else {
    normalizedAnswer = currentAnswer.trim();
    normalizedDefault = default_answer.trim();
  }

  // Compare normalized values
  if (op === 'NEQ_DEFAULT')
    return !_.isEqual(normalizedAnswer, normalizedDefault);
  return false;
};

const evaluateLiteral = (guard, answers) => {
  const answer = answers[guard.question_id];
  // string or undefined case; should never happen actually
  return Array.isArray(answer) && answer.includes(guard.answer_id);
};

const evaluateGuard = (guard, answers, questions) => {
  if (guard === null) {
    return true;
  }
  // Besides others, free text questions all have null guards and will enter the
  // above conditional. So cast `answers` according to the type expected by `evaluateLiteral`
  if (!Array.isArray(guard)) {
    if ('op' in guard) {
      return evaluateDefaultCheck(guard, answers, questions);
    } else {
      return evaluateLiteral(guard, answers);
    }
  }
  const [operator, ...args] = guard;
  switch (operator) {
    case 'AND':
      return args.every((arg) => evaluateGuard(arg, answers, questions));
    case 'OR':
      return args.some((arg) => evaluateGuard(arg, answers, questions));
    default: {
      // TypeScript pattern to ensure exhaustive checking of operators
      const _exhaustiveCheck = operator;
      return _exhaustiveCheck;
    }
  }
};

const findNextQuestionId = (flows, currentAnswers, questions) => {
  if (!flows.length) return null;

  const defaultFlow = flows.find((flow) => !flow.guard);
  const guardedFlows = flows.filter((flow) => flow.guard);

  // Try guarded flows first
  const matchingGuardedFlow = guardedFlows
    .filter((flow) => evaluateGuard(flow.guard, currentAnswers, questions))
    .sort((f0, f1) => f1.priority - f0.priority)[0];

  if (matchingGuardedFlow) {
    return matchingGuardedFlow.target_question_id;
  }

  return defaultFlow?.target_question_id ?? null;
};

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
  /**
   * @param {object} [options={}] - Optional parameters for customization.
   * @param {object} [options.paragraphOptions={}] - Custom styles for the answer's Paragraph.
   * @param {object} [options.textOptions={}] - Custom styles for the answer's TextRun.
   * @param {object} [options.selectedTextOptions={}] - Custom styles for selected answers' TextRun(s).
   */
  const showAllAnswersToChoiceQuestion = (
    serial,
    assessment,
    questionnaire,
    options = {}
  ) => {
    const {
      paragraphOptions = {},
      textOptions = {},
      selectedTextOptions = { bold: true, underline: true, ...textOptions },
    } = options;

    const questionId = findQuestionIdBySerial(serial, questionnaire);
    const possibleAnswers = questionnaire[questionId].answers;
    const selectedAnswerIds =
      assessment.content.questionnaireProgress.answers[questionId];

    return possibleAnswers.map(({ id, content }, index) => {
      const isSelectedAnswer = selectedAnswerIds.includes(id);

      return new docx.Paragraph({
        bullet: { level: 0 },
        spacing: { after: index === possibleAnswers.length - 1 ? 0 : 50 },
        ...paragraphOptions,
        children: [
          new docx.TextRun({
            text: content,
            // Use textOptions for unselected, and the more specific
            // selectedTextOptions for selected items.
            ...(isSelectedAnswer ? selectedTextOptions : textOptions),
          }),
        ],
      });
    });
  };

  /**
   * @param {object} [options={}] - Optional parameters for customization.
   * @param {object} [options.paragraphOptions={}] - Custom styles for the answer's Paragraph.
   * @param {object} [options.textOptions={}] - Custom styles for the answer's TextRun.
   */
  const showAnswer = (serial, assessment, questionnaire, options = {}) => {
    const { paragraphOptions = {}, textOptions = {} } = options;

    const questionId = findQuestionIdBySerial(serial, questionnaire);
    if (!questionId) return [];

    const question = questionnaire[questionId];
    const answerData =
      assessment.content.questionnaireProgress.answers[questionId];
    if (!answerData) return [];

    if (question.type !== 'FREE_TEXT') {
      const selectedAnswerIds = Array.isArray(answerData)
        ? answerData
        : [answerData];
      const selectedAnswers = question.answers.filter((possibleAnswer) =>
        selectedAnswerIds.includes(possibleAnswer.id)
      );
      return selectedAnswers.map(
        ({ content }) =>
          new docx.Paragraph({
            ...paragraphOptions,
            children: [new docx.TextRun({ text: content, ...textOptions })],
          })
      );
    }

    return [
      new docx.Paragraph({
        ...paragraphOptions,
        children: [new docx.TextRun({ text: answerData, ...textOptions })],
      }),
    ];
  };

  const showAnswerConditionally = (
    serial,
    assessment,
    questionnaire,
    hidingConditions,
    title
  ) => {
    const questionId = findQuestionIdBySerial(serial, questionnaire);
    const allAnswers = assessment.content.questionnaireProgress.answers;
    const userAnswers = allAnswers[questionId];
    // NOTE assuming serial does not refer to a free text question & userAnswers is a list
    const answersData = userAnswers.map((answerId) =>
      questionnaire[questionId].answers.find(({ id }) => id === answerId)
    );
    const answerContents = answersData.map(({ content }) => content);
    if (hidingConditions.some((cond) => answerContents.includes(cond))) {
      return [];
    }
    const dependentQuestionId = findNextQuestionId(
      questionnaire[questionId].flows,
      allAnswers,
      questionnaire
    );
    const dependentQuestion = questionnaire[dependentQuestionId];
    return [
      new docx.Paragraph({
        children: [new docx.TextRun({ text: title, bold: true })],
        spacing: { before: 200, after: 200 },
      }),
      ...showAnswer(dependentQuestion.serial, assessment, questionnaire),
    ];
  };

  /**
   * Creates a self-contained block with a formatted title and its corresponding answer.
   *
   * @param {string} serial - The question serial number used to find the answer.
   * @param {string} title - The text to display as the question's title.
   * @param {object} assessment - The assessment object.
   * @param {object} questionnaire - The questionnaire object.
   * @param {object} [options={}] - Optional parameters for customization.
   * @param {'default' | 'allOptions'} [options.displayType='default'] - The type of answer display.
   * @param {object} [options.titleParagraphOptions={}] - Custom styles for the title's Paragraph.
   * @param {object} [options.titleTextOptions={}] - Custom styles for the title's TextRun.
   * @param {object} [options.answerParagraphOptions={}] - Custom styles for the answer's Paragraph(s).
   * @param {object} [options.answerTextOptions={}] - Custom styles for the answer's TextRun(s).
   * @returns {Array<docx.Paragraph>} An array of docx objects for the block.
   */
  const createQuestionBlock = (
    serial,
    title,
    assessment,
    questionnaire,
    options = {}
  ) => {
    const {
      displayType = 'default',
      titleParagraphOptions = {},
      titleTextOptions = {},
      answerParagraphOptions = {},
      answerTextOptions = {},
      answerSelectedTextOptions,
    } = options;

    const titleParagraph = new docx.Paragraph({
      spacing: { before: 200, after: 200 },
      ...titleParagraphOptions,
      children: [
        new docx.TextRun({ text: title, bold: true, ...titleTextOptions }),
      ],
    });

    const answerOptions = {
      paragraphOptions: answerParagraphOptions,
      textOptions: answerTextOptions,
      ...(answerSelectedTextOptions && {
        selectedTextOptions: answerSelectedTextOptions,
      }),
    };

    const answerFunction =
      displayType === 'allOptions'
        ? showAllAnswersToChoiceQuestion
        : showAnswer;
    const answerParagraphs = answerFunction(
      serial,
      assessment,
      questionnaire,
      answerOptions
    );

    return [titleParagraph, ...answerParagraphs];
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
      ...createQuestionBlock(
        '3.1',
        'Vulnerability Assessment for',
        assessment,
        questionnaire,
        {
          titleParagraphOptions: {
            spacing: { before: 800 },
            alignment: docx.AlignmentType.CENTER,
          },
          titleTextOptions: { size: 36, italics: true, bold: false },
          answerParagraphOptions: {
            spacing: { before: 800 },
            alignment: docx.AlignmentType.CENTER,
          },
          answerTextOptions: { size: 44 },
        }
      ),
      ...createQuestionBlock('2.1.1', 'located at', assessment, questionnaire, {
        titleParagraphOptions: {
          spacing: { before: 800 },
          alignment: docx.AlignmentType.CENTER,
        },
        titleTextOptions: { size: 36, italics: true },
        answerParagraphOptions: {
          spacing: { before: 800 },
          alignment: docx.AlignmentType.CENTER,
        },
        answerTextOptions: { size: 44 },
      }),
      ...createQuestionBlock('3.6.1', 'on date(s)', assessment, questionnaire, {
        titleParagraphOptions: {
          spacing: { before: 800 },
          alignment: docx.AlignmentType.CENTER,
        },
        titleTextOptions: { size: 36, italics: true },
        answerParagraphOptions: {
          spacing: { before: 800 },
          alignment: docx.AlignmentType.CENTER,
        },
        answerTextOptions: { size: 44 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: `Assessment date: ${new Date().toISOString()}`,
            size: 28,
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

    const cellBorders = {
      top: { style: docx.BorderStyle.SINGLE, size: 6, color: 'auto' },
      bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: 'auto' },
      left: { style: docx.BorderStyle.SINGLE, size: 6, color: 'auto' },
      right: { style: docx.BorderStyle.SINGLE, size: 6, color: 'auto' },
    };

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
      ...createQuestionBlock(
        '2.1.1',
        'Description of the venue location',
        assessment,
        questionnaire,
        {
          titleParagraphOptions: { spacing: { after: 200 } },
        }
      ),
      ...createQuestionBlock(
        '2.2',
        'Sectors applicable for this location',
        assessment,
        questionnaire,
        { displayType: 'allOptions' }
      ),
      ...createQuestionBlock(
        '2.7',
        'Venue location type',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '2.7.1',
        'Venue type characteristics',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '2.13',
        'Overall size (gross area)',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '2.10',
        'Characteristics of surrounding area',
        assessment,
        questionnaire
      ),
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
      new docx.Paragraph({
        text: 'Event Information',
        heading: docx.HeadingLevel.HEADING_1,
        spacing: { before: 200, after: 200 },
      }),
      ...createQuestionBlock(
        '3.2',
        'Main activities (program/agenda)',
        assessment,
        questionnaire,
        { titleParagraphOptions: { spacing: { after: 200 } } }
      ),
      ...createQuestionBlock('3.5', 'Frequency', assessment, questionnaire),
      ...createQuestionBlock('3.6', 'Duration', assessment, questionnaire),
      ...createQuestionBlock(
        '3.8',
        'Event opening and closing times for the public (day/time)',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '3.9',
        'Maximum expected attendance',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '3.11',
        'Expected average crowd density (participants/m²)',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '3.16',
        'At what level is the event known?',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '3.17',
        "Activity/event's character",
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '3.19',
        'Media coverage and live broadcasting',
        assessment,
        questionnaire
      ),
    ];
  };

  const createThreatInformation = (assessment, questionnaire) => {
    return [
      new docx.Paragraph({
        text: 'Threat-related Information',
        heading: docx.HeadingLevel.HEADING_1,
        spacing: { before: 200, after: 200 },
      }),
      ...showAnswerConditionally(
        '4.1.1',
        assessment,
        questionnaire,
        ['No such categorization exists'],
        'Current national terrorism threat level'
      ),
      ...createQuestionBlock(
        '4.1.2',
        'Availability of information on completed/foiled/planned attacks and detail of provided analysis from national authorities',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.1.3',
        'Availability of information on active terrorist groups and relevant trends (e.g. political of ideology motivated) from national authorities',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.1.4',
        'Availability of information on local threat level from authorities',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.1.5',
        'Local threat level at the moment of the assessment',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.3.2',
        'Specific threats/public statements by potential aggressors against civil targets or entities directly connected to the venue/event activity (e.g. owners/operators, suppliers)?',
        assessment,
        questionnaire
      ),
      ...showAnswerConditionally(
        '4.3.2',
        assessment,
        questionnaire,
        ['No', 'Not known'],
        'When were these specific threats/public statements made?'
      ),
      ...createQuestionBlock(
        '4.3.3',
        'Have similar civil targets been subjected to attacks in the past?',
        assessment,
        questionnaire
      ),
      ...showAnswerConditionally(
        '4.3.3',
        assessment,
        questionnaire,
        ['No', 'Not known'],
        'When did the most recent attacks(s) happen?'
      ),
      ...showAnswerConditionally(
        '4.3.4',
        assessment,
        questionnaire,
        ['No', 'Not known'],
        'Previous incidents and suspicious activities at venue location'
      ),
      ...createQuestionBlock(
        '4.3.5',
        'Has the venue location undergone a specific threat assessment in the past and is it available?',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.3.6',
        "Venue's notoriety",
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.3.7',
        'Iconic status (symbolic/historical/cultural value) of the venue location',
        assessment,
        questionnaire
      ),

      ...createQuestionBlock(
        '4.3.8',
        'Presence of VIPs of infamous/controversial entities at the activity/event?',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.3.9',
        "Does the event itself of the associated activity represent a religious/ethno-nationalist/extremism ideology that may be considered 'contrary and/or offensive' according to certain belief systems or ideological/moral agendas?",
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.3.11',
        'What type of threat actor may be particularly interested in the venue location and its activities/events?',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.3.11.2',
        'Main motivations for the choice of this type of threat actor',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.3.12',
        'Modus operandi considered most credible for an actual attack and therefore considered a priority?',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.3.12.2',
        'Main motivations for this choice of this type of modus operandi',
        assessment,
        questionnaire
      ),
      ...createQuestionBlock(
        '4.4.2',
        'Justification for the exclusion of standard threat scenarios from the scope of this assessment',
        assessment,
        questionnaire
      ),
    ];
  };

  const newPage = () =>
    new docx.Paragraph({ children: [new docx.PageBreak()] });

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
          ...createThreatInformation(ASSESSMENT, FLAT_QUESTIONNAIRE),
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
