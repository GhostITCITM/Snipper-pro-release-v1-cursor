import { createWorker, Worker } from "tesseract.js";

let _worker: Worker | null = null;

async function getWorker(): Promise<Worker> {
  if (_worker) return _worker;

  _worker = await createWorker({
    logger: (m) => console.log("OCR:", m)
  });

  await _worker.loadLanguage("eng");
  await _worker.initialize("eng");

  await _worker.setParameters({
    tessedit_char_whitelist:
      "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz .,;:!?@#$%^&*()_+-=[]{}|\\;':\",./<>?`~",
    tessedit_pageseg_mode: "6"
  });

  return _worker;
}

export async function recognizeText(img: ImageData): Promise<string> {
  const worker = await getWorker();
  const { data } = await worker.recognize(img);
  return data.text.trim();
}

export async function recognizeNumbers(imageData: ImageData): Promise<number[]> {
  const text = await recognizeText(imageData);
  const numberRegex = /-?\d+(?:[.,]\d+)?/g;
  const matches = text.match(numberRegex) || [];

  return matches
    .map((match) => {
      const normalized = match.replace(/,/g, "");
      return parseFloat(normalized);
    })
    .filter((num) => !isNaN(num));
}

export class OCREngine {
  static async initialize(): Promise<void> {
    await getWorker();
  }

  static async terminate(): Promise<void> {
    if (_worker) {
      await _worker.terminate();
      _worker = null;
    }
  }

  static async recognizeText(imageData: ImageData): Promise<{ text: string; confidence: number }> {
    const worker = await getWorker();
    const { data } = await worker.recognize(imageData);
    return {
      text: data.text.trim(),
      confidence: data.confidence
    };
  }

  static async recognizeNumbers(imageData: ImageData): Promise<number[]> {
    return recognizeNumbers(imageData);
  }
}

export class SumCalculator {
  static calculateSum(numbers: number[]): number {
    return numbers.reduce((sum, num) => sum + num, 0);
  }

  static async extractAndSum(imageData: ImageData): Promise<{
    sum: number;
    numbers: number[];
    text: string;
  }> {
    const numbers = await recognizeNumbers(imageData);
    const sum = this.calculateSum(numbers);
    const text = await recognizeText(imageData);

    return {
      sum,
      numbers,
      text
    };
  }
}
