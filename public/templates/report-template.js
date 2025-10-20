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
  const newPage = () =>
    new docx.Paragraph({ children: [new docx.PageBreak()] });

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
        titleTextOptions: { size: 36, italics: true, bold: false },
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
        titleTextOptions: { size: 36, italics: true, bold: false },
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

  const createThreatRating = (images, useBarCharts = false) => {
    const chartB64 = useBarCharts
      ? images.threatRatingBarChart
      : images.threatRatingRadarChart;

    const threatRatingChart = chartB64
      ? new docx.Paragraph({
          children: [
            new docx.ImageRun({
              data: chartB64.replace('data:image/png;base64,', ''),
              type: 'png',
              transformation: { width: 400, height: 400 },
            }),
          ],
          spacing: { before: 200, after: 200 },
        })
      : new docx.Paragraph({
          text: 'THERE SHOULD BE AN IMAGE HERE',
          spacing: { after: 200 },
        });
    return [
      new docx.Paragraph({
        text: 'Threat Rating',
        heading: docx.HeadingLevel.HEADING_1,
        spacing: { before: 200, after: 200 },
      }),

      new docx.Paragraph({
        text: "The Threat Rating in this assessment considers the availability of national & local threat information sources, the public's awareness regarding terrorism-related threats and eventual preparedness guidance. It also evaluates the venue and event-specific threat history.",
        spacing: { after: 200 },
      }),

      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Results for the assessed venue and event/activity:',
            bold: true,
          }),
        ],
        spacing: { after: 400 }, // Added extra spacing for the plot area
      }),

      threatRatingChart,

      new docx.Paragraph({
        text: 'The Threat Rating is presented as a factor ranging from 0 ‘Very Low’ (Dark Green) to 1 ‘Very High’ (Dark Red). It is shown as a continuous scale. Colours are associated with a score range and level.',
        spacing: { after: 200 },
      }),

      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'The ',
          }),
          new docx.TextRun({
            text: 'overall Threat Rating T',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' is composed of four main factors, representing: Operational capability T',
          }),
          new docx.TextRun({
            text: 'oc',
            subScript: true,
          }),
          new docx.TextRun({
            text: ', Intention & History T',
          }),
          new docx.TextRun({
            text: 'ih',
            subScript: true,
          }),
          new docx.TextRun({
            text: ', Activity T',
          }),
          new docx.TextRun({
            text: 'ac',
            subScript: true,
          }),
          new docx.TextRun({
            text: ' and Operating Environment T',
          }),
          new docx.TextRun({
            text: 'oe',
            subScript: true,
          }),
          new docx.TextRun({
            text: '.',
          }),
        ],
        spacing: { after: 400 }, // Added extra spacing for the table area
      }),

      // --- You can insert your rating table here ---

      new docx.Paragraph({
        text: 'All factors are equally weighed. They represent:',
        spacing: { after: 200 },
      }),

      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Operational Capability',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' is the acquired, assessed, or demonstrated level of operational capability to conduct terrorist attacks. It includes the required expertise & technical training, associated resources & costs (weapons, manpower, material), as well as the difficulty to obtain such resources.',
          }),
        ],
        bullet: { level: 0 },
        spacing: { after: 100 },
      }),

      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Intention & History',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' represents two closely related factors. Intention is the motivation of terrorists to conduct an attack. History addresses the fact that previous attempts are potentially good indicators of future attempts. It also reflects the fact that local or regional and recent attempts are potentially better indicators.',
          }),
        ],
        bullet: { level: 0 },
        spacing: { after: 100 },
      }),

      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Activity',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' considers the level of terrorist activity in a country. Such activity may not always present a threat to local interests as terrorists may use countries as support bases and may not want to jeopardize their status by conducting terrorist acts there.',
          }),
        ],
        bullet: { level: 0 },
        spacing: { after: 100 },
      }),

      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Operating Environment',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' considers how the overall environment, including political and security considerations, influences the ability and motivation of terrorists to conduct an attack.',
          }),
        ],
        bullet: { level: 0 },
        spacing: { after: 200 },
      }),

      new docx.Paragraph({
        text: 'Questions in the assessment process may influence only one or several of these factors.',
      }),
    ];
  };

  const createVulnerabilityRatingSection = (images, useBarCharts = false) => {
    const chartB64 = useBarCharts
      ? images.vulnerabilityRatingBarChart
      : images.vulnerabilityRatingRadarChart;

    const vulnerabilityRatingChart = chartB64
      ? new docx.Paragraph({
          children: [
            new docx.ImageRun({
              data: chartB64.replace('data:image/png;base64,', ''),
              type: 'png',
              transformation: { width: 400, height: 400 },
            }),
          ],
          spacing: { before: 200, after: 200 },
        })
      : new docx.Paragraph({
          text: 'THERE SHOULD BE AN IMAGE HERE',
          spacing: { after: 200 },
        });

    return [
      new docx.Paragraph({
        text: 'Vulnerability Rating',
        heading: docx.HeadingLevel.HEADING_1,
        spacing: { before: 200, after: 200 },
      }),
      new docx.Paragraph({
        text: 'The Vulnerability Rating in this assessment focuses on analysing the vulnerabilities of the venue location and the related activity/event. It considers potential vulnerabilities through a structured approach based on best practice and relevant criteria. Vulnerabilities may be related to certain areas/sectors of the venue location, to a specific attack type/modus operandi or to deficiencies in the human, technical and organizational domain.',
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Vulnerability Rating results for the assessed venue and event/activity:',
            bold: true,
          }),
        ],
        spacing: { after: 400 }, // Added extra spacing for the plot area
      }),
      vulnerabilityRatingChart,
      new docx.Paragraph({
        text: 'The Vulnerability Rating is presented as a factor ranging from 0 ‘Very Low’ (Green) to 1 ‘Very High’ (Red). It is shown as a continuous scale. Colours are associated with a score range and level.',
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun('The '),
          new docx.TextRun({
            text: 'overall Vulnerability Rating V',
            bold: true,
            italics: true,
          }),
          new docx.TextRun(
            ' is composed of ten factors, representing: Importance V'
          ),
          new docx.TextRun({ text: 'im', subScript: true }),
          new docx.TextRun(', Business continuity - Resilience V'),
          new docx.TextRun({ text: 'bcr', subScript: true }),
          new docx.TextRun(', Reputational damage V'),
          new docx.TextRun({ text: 'rd', subScript: true }),
          new docx.TextRun(', Venue Location V'),
          new docx.TextRun({ text: 'lc', subScript: true }),
          new docx.TextRun(', Symbolism V'),
          new docx.TextRun({ text: 'sy', subScript: true }),
          new docx.TextRun(', Accessibility V'),
          new docx.TextRun({ text: 'ac', subScript: true }),
          new docx.TextRun(', Uniqueness V'),
          new docx.TextRun({ text: 'un', subScript: true }),
          new docx.TextRun(', Attendance & Distribution V'),
          new docx.TextRun({ text: 'ad', subScript: true }),
          new docx.TextRun(', Existing measures V'),
          new docx.TextRun({ text: 'em', subScript: true }),
          new docx.TextRun(' and Media attention V'),
          new docx.TextRun({ text: 'ma', subScript: true }),
          new docx.TextRun('.'),
        ],
      }),
      new docx.Paragraph({
        text: 'All factors are equally weighed. They consider the following aspects:',
        spacing: { before: 200, after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: 'Importance', bold: true, italics: true }),
          new docx.TextRun({
            text: ' depends on the public space’s functions, its interdependencies with other facilities and the collateral consequences for the state and the society of a potential attack.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Business continuity, Resilience',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' is a factor taking into account the amount of time required to continue (in a degraded way) or re-establish operations and activities. This can be either temporary or permanent.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Reputational damage',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' accounts for the perception of reputational repercussions associated with a potential attack. Considerations could include adverse publicity, erosion of confidence, and the perception of poor security. Note that reputational damage is to a certain degree dependent on media attention.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Venue Location.',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' This factor reflects the assumption that events and activities depending on their physical location are more likely to be the targets of an attack and that for example the threat is higher near major population centres.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: 'Symbolism', bold: true, italics: true }),
          new docx.TextRun({
            text: ' is linked to the prominence, iconic value and attractiveness of a public space as a potential target and its probability of being considered as promoting a lifestyle that is against the political, social or religious ideology of attackers. Popular tourist locations, landmarks and cultural sites but also spaces that may be associated with ethnic or racial minorities. It accounts for the fact that some public spaces are more controversial and well known. This factor is based on the assumption that public spaces that have a high symbolism value are more likely targets.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Accessibility',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' is a measure of the venue’s ‘openness’ and how difficult it would be for an aggressor to enter. This factor addresses the degree of a controlled perimeter or access control.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: 'Uniqueness', bold: true, italics: true }),
          new docx.TextRun({
            text: ' is a factor that assesses how common the type of public space or venue is within the surrounding environment. It reflects the greater likelihood that an aggressor will attempt to target a particular location if it is the only one of its kind (its uniqueness).',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Attendance & Distribution',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: " accounts for the maximum number of people (personnel and visitors) present in the public space, their importance (VIP presence) and the associated variation over time and space (throughout the venue site). This considers the likelihood that an aggressor tries to maximize the attack's impact by targeting locations with high crowd density and/or presence of VIPs.",
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Existing measures',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' considers security measures that are already present in the examined public space or venue and may render it less attractive to attackers and/or the presence of vulnerabilities that make it more appealing to aggressors.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Media attention',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' focuses on the publicity the aggressor could expect if targeting a specific venue location and event/activity.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        text: 'Questions in the assessment process may influence only one or several of these factors.',
        spacing: { before: 200 },
      }),
    ];
  };

  const createHtoSection = (images, useBarCharts = false) => {
    const chartB64 = useBarCharts
      ? images.vulnerabilityRatingBarChart
      : images.vulnerabilityRatingRadarChart;

    const htoChart = chartB64
      ? new docx.Paragraph({
          children: [
            new docx.ImageRun({
              data: chartB64.replace('data:image/png;base64,', ''),
              type: 'png',
              transformation: { width: 400, height: 400 },
            }),
          ],
          spacing: { before: 200, after: 200 },
        })
      : new docx.Paragraph({
          text: 'THERE SHOULD BE AN IMAGE HERE',
          spacing: { after: 200 },
        });

    return [
      new docx.Paragraph({
        text: 'HTO concept assessing the comprehensiveness of measures:',
        heading: docx.HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 },
      }),
      new docx.Paragraph({
        text: 'The HTO (Human, Technical, Organizational) concept is commonly applied to analyse complex activities. In this context, a comprehensive security setup should be composed of all three components. In the HTO concept:',
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'The “Human (H)”',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' stands for aspects that are individual and at the same time important to perform a task or a change. Such aspects may, for example, include individual skill, knowledge, experiences or established relations with other people.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'The “Technical (T)”',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' stands for technical system and can be divided into two parts. A technical system can be evaluated in relation to technical limitations, problems (both recurrent and stochastic), availability, and reliability. Technical systems can also assist decision-makers as illustrated by information systems, hardware and software used as support tools.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'The “Organizational (O)“',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' component comprises how an activity is organized and structured. Examples are responsibilities and powers, policies and strategies. This also includes rules, procedures and other factors, whether formal or informal.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        text: 'The report presents a separate vulnerability rating for every one of these factors. It may provide an orientation where to direct efforts for example for the implementation of new, complementary measures.',
        spacing: { before: 200, after: 200 },
      }),
      htoChart,
    ];
  };

  const createAttackPhaseSection = () => {
    return [
      new docx.Paragraph({
        text: 'Vulnerability ratings associated with distinctive attack phases',
        heading: docx.HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 },
      }),
      new docx.Paragraph({
        text: 'In asset protection, measures have often a specific function that refers to distinctive phases prior or during an attack. These are:',
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Prepare, Prevent and Deter:',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' A good state of preparedness is generally associated with a high level of resilience and a low level of vulnerability. Certain aspects comprising the preparedness level are visible to a potential aggressor, having a deterrence function. Therefore, this factor influences the target selection process, as successful deterrence may eventually avoid an attack.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: 'Detect:', bold: true, italics: true }),
          new docx.TextRun({
            text: ' If an attack occurs, early detection is a crucial factor to limit consequences and be able to initiate an appropriate response. Detection mostly occurs through human observation and/or technical systems.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Alert and Assess:',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' To initiate an appropriate response, the detection is transformed into an alert and a process that assesses the severity of the incident. If deemed credible and serious, it initiates the response.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: 'Delay:', bold: true, italics: true }),
          new docx.TextRun({
            text: ' The progression of aggressors can be hindered through delay measures, thereby gaining valuable time for intervention forces to respond.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Protect and Deny:',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' In some cases, effective protective measures can deny aggressor access. This can be the case for a hardened building or site perimeter, for example a bollard stopping a hostile vehicle.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Respond and Recover:',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' The response capability determines the ability to defend against or eventually defeat an aggressor. Following the containment or elimination of the threat, recovery actions initiate to re-establish ‘normal’ functioning.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        text: 'The report presents a separate vulnerability rating for every one of these phases. Some phases may be unsuitable or difficult to implement for certain venue locations, especially those with a high degree of ‘openness’. Nevertheless, this evaluation may provide an orientation for the direction of reinforcement efforts.',
        spacing: { before: 200, after: 200 },
      }),

      // attackPhasesChart,
    ];
  };

  const createAttackFocusSection = () => {
    return [
      new docx.Paragraph({
        text: 'Assessment focus according to attack phases',
        heading: docx.HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 200 },
      }),
      new docx.Paragraph({
        text: 'Vulnerabilities may also be associated with the aggressor’s attack phases. These are:',
        spacing: { after: 200 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Target selection:',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' This phase is related to the target selection process before an attack. All factors presented in the overall Vulnerability rating V presented above may contribute to this process.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Planning and Dry run:',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' Following the identification of a suitable target, the planning phase of the attack may involve observing, testing and probing existing security measures. Some of these activities may be detectable, potentially disrupting an attack before it occurs.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({ text: 'Execution:', bold: true, italics: true }),
          new docx.TextRun({
            text: ' During the execution phase of an attack, the objective of an aggressor is to maximize the impact. Effective countermeasures can limit this impact.',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: 'Escape and exploitation:',
            bold: true,
            italics: true,
          }),
          new docx.TextRun({
            text: ' Escape in the aftermath of an attack may be of importance to aggressors (this for example is not the case in any ‘suicide attack’). This evaluation focuses on the aggressor’s objective to exploit and provide visibility to the attack. Media attention and live broadcasting are therefore important factors (even if self-streaming by aggressors over the internet is becoming a more common feature).',
          }),
        ],
        bullet: { level: 0 },
      }),
      new docx.Paragraph({
        text: 'Hereunder, the report does not evaluate a degree of vulnerability but presents an overview to what extent the assessment focused on different adversary attack phases.',
        spacing: { before: 200, after: 200 },
      }),

      // Placeholder for the Attack Phases pie chart
      // attackPhasesPieChart,
    ];
  };

  const createVulnerabilityRating = (images, useBarCharts = false) => {
    return [
      ...createVulnerabilityRatingSection(images, useBarCharts),
      ...createHtoSection(images, useBarCharts),
      ...createAttackPhaseSection(),
      ...createAttackFocusSection(),
    ];
  };

  const createHighlightedRun = (text) => {
    return new docx.TextRun({
      text: text,
      color: 'FFFFFF', // White text
      shading: {
        type: docx.ShadingType.SOLID,
        fill: '#003399', // Blue color
      },
    });
  };

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
          createHighlightedRun(2),
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
          ...createThreatRating({}),
          newPage(),
          ...createVulnerabilityRating({}),
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
