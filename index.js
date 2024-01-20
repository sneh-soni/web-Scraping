import fetch from "node-fetch";
import { createClient } from "pexels";
import * as XLSX from "xlsx";

const searchQuery = [
  "Products",
  "advertisement",
  "digital marketing",
  "product photography",
  "product advertising",
];

const client = createClient(
  "7VuVNqX4UWteorX8Y1PY0tPKdtQjUzrjZd4WC4pDM05qXaTDKYQuvm8S"
);
const adobeApiKey = "5120837ff0544503a282ef91a27037e2";
const adobeProduct = "Project 1";
const adobeLocale = "en_US";

const searchPhotos = async (query) => {
  try {
    const photos = await client.photos.search({ query, per_page: 80 });
    return photos.photos.map((photo) => ({
      id: photo.id,
      type: "image/jpeg",
      height: photo.height + " px",
      width: photo.width + " px",
      alt: photo.alt,
      url: photo.url,
    }));
  } catch (error) {
    console.error(
      `Error searching photos for query '${query}': ${error.message}`
    );
    return [];
  }
};

const searchVideos = async (query) => {
  try {
    const videos = await client.videos.search({ query, per_page: 80 });
    return videos.videos.map((video) => ({
      id: video.id,
      type: "video/quicktime",
      height: video.height + " px",
      width: video.width + " px",
      alt: "",
      url: video.url,
    }));
  } catch (error) {
    console.error(
      `Error searching videos for query '${query}': ${error.message}`
    );
    return [];
  }
};

const adobeSearch = async (query) => {
  const url = `https://stock.adobe.io/Rest/Media/1/Search/Files?locale=${adobeLocale}&search_parameters[words]=${query}`;

  try {
    const response = await fetch(url, {
      headers: {
        "x-api-key": adobeApiKey,
        "x-product": adobeProduct,
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    const updatedData = data.files.filter(
      (file) => file.content_type !== "application/illustrator"
    );

    return updatedData.map((file) => ({
      id: file.id,
      type: file.content_type,
      height: file.thumbnail_height + " px",
      width: file.thumbnail_width + " px",
      alt: file.title,
      url: file.thumbnail_url,
    }));
  } catch (error) {
    console.error(`Error fetching data for query '${query}': ${error.message}`);
    return [];
  }
};

const createExcelSheet = (data) => {
  const workbook = XLSX.utils.book_new();
  const mergedData = data.reduce((acc, result) => [...acc, ...result], []);
  const sheet = XLSX.utils.json_to_sheet(mergedData);
  XLSX.utils.book_append_sheet(workbook, sheet, "Merged Results");
  XLSX.writeFile(workbook, "results.xlsx");
};

const executeSearches = async () => {
  const imageResults = await Promise.all(
    searchQuery.map((query) => searchPhotos(query))
  );
  const videoResults = await Promise.all(
    searchQuery.map((query) => searchVideos(query))
  );
  const adobeResults = await Promise.all(
    searchQuery.map(async (query) => adobeSearch(query))
  );

  const combinedResults = [
    imageResults.flat(),
    videoResults.flat(),
    adobeResults.flat(),
  ];

  createExcelSheet(combinedResults);
};

executeSearches();
