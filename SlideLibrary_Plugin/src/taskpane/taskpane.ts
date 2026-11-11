import "./powerpoint";
import { getPresentationBlob } from "../office/presentation";

async function handleClick() {
  try {
    const { blob, fileName } = await getPresentationBlob();
    console.log("Файл успешно собран!", fileName, "Размер в байтах:", blob.size);
  } catch (error) {
    console.error("Ошибка при чтении презентации:", error);
  }
}
