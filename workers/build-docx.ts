/**
 * The worker receives a message with the user's code and returns a message with the Blob or an error.
 * We use a Web Worker to run the user's code in a separate thread and avoid blocking the main thread.
 */
import { buildDocx } from '@/lib/build-docx';

// 👇 1. Define the shape of the incoming assets
interface AssetData {
  assessment: object;
  questionnaire: object;
  commissionBanner: string;
  unitLogo: string;
}

self.onmessage = (
  event: MessageEvent<{ name: string; text: string; assets: AssetData }>
) => {
  const {
    name,
    text,
    assets: { questionnaire, assessment, commissionBanner, unitLogo },
  } = event.data;
  buildDocx(text, assessment, questionnaire, commissionBanner, unitLogo)
    .then((blob) => {
      self.postMessage({
        status: 'success',
        payload: {
          name,
          text,
          blob,
          buildError: undefined,
        },
      });
    })
    .catch((error) => {
      self.postMessage({
        status: 'error',
        payload: {
          name,
          text,
          blob: undefined,
          buildError: error,
        },
      });
    });
};
