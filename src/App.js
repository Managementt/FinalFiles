import React, { useState } from "react";
import * as XLSX from "xlsx";
import AWS from "aws-sdk";
import './styles.css';

// Configure Wasabi storage
AWS.config.update({
  accessKeyId: '3MF2DSVWVGOI3S47VKY2',
  secretAccessKey: '65GzKkrWaBMgPFrMBtQrPEeGYAHIFMMTCj29QiqB',
  region: "ca-central-1",
});

const s3 = new AWS.S3({
  endpoint: 'https://s3.ca-central-1.wasabisys.com',  // Update with your Wasabi region
  s3ForcePathStyle: true,
});

const UploadForm = () => {
  const [fileType, setFileType] = useState("");
  const [formData, setFormData] = useState({});

  const handleChange2 = (e) => {
    const { name, value, type, files } = e.target;
  
    if (type === "file") {
      // Convert the FileList to an array
      const newFiles = Array.from(files);
  
      setFormData((prevData) => {
        // Ensure prevData.files is always an array
        const existingFiles = Array.isArray(prevData.files) ? prevData.files : [];
  
        // Remove the file that was replaced (if any) and only keep the current files
        const updatedFiles = existingFiles.filter((file) => 
          !newFiles.some((newFile) => newFile.name === file.name)
        );
  
        // Combine the remaining files with the new files
        return {
          ...prevData,
          files: [...updatedFiles, ...newFiles], // Append new files while removing old ones
        };
      });
    } else {
      setFormData((prevData) => ({
        ...prevData,
        [name]: value,
      }));
      console.log(formData)
    }
  };

  const uploadToWasabi2 = async (file,bin) => {
    try {
      console.log(bin)
      const fileName = `${bin}/${file.name}`; 
      const uploadParams = {
        Bucket: "may2025", // Target bucket
        Key: fileName, // Folder path + file name
        Body: file,
        ACL: "public-read",
        ContentType: file.type,
      };
  
      const response = await s3.upload(uploadParams).promise();
      console.log("Uploaded file URL:", response.Location); // Debugging
      return response.Location; // Return file URL
    } catch (error) {
      console.error("Error uploading file:", error);
      return "";
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    let wasabiUrls = [];
    // Upload multiple files from the form
    for (const file of formData.files || []) {
      if (file) {
        console.log(formData.bin)
        const url = await uploadToWasabi2(file, formData.bin);
        wasabiUrls.push({ fileName: file.name, url });
      }
    }

    let wasabiUrl = "";
    if (formData.uploadedFile) {
      wasabiUrl = await uploadToWasabi2(formData.uploadedFile, formData.bin);
    }

    // Define the file name where data will be stored
    const fileName = "form_data.xlsx";

    // Try to fetch the existing Excel file from Wasabi
    let existingData = [];
    try {
      const params = { Bucket: "excelfile", Key: fileName };
      const existingFile = await s3.getObject(params).promise();
      
      // Read the existing Excel file
      const workbook = XLSX.read(existingFile.Body, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      existingData = XLSX.utils.sheet_to_json(sheet);
    } catch (error) {
      console.log("No existing file found. Creating a new one.");
    }

    // Append new form data
    const newEntry = {
      Timestamp: new Date().toISOString(),
      FileType: fileType,
      FileStatus: formData.fileStatus || "",
      BookDetails: formData.bookDetails || "",
      BIN: formData.bin || "",
      WasabiFileURL: wasabiUrl || "",
    };

    existingData.push(newEntry);

    // Convert JSON to worksheet
    const ws = XLSX.utils.json_to_sheet(existingData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Form Data");

    // Write the updated Excel file to buffer
    const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const updatedFile = new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    // Upload updated file back to Wasabi
    const uploadParams = {
      Bucket: "excelfile",
      Key: fileName,
      Body: updatedFile,
      ACL: "public-read",
      ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    };

    try {
      await s3.upload(uploadParams).promise();
      alert("Form data successfully uploaded!");
      setFormData({}); // Clear the form data state
      setFileType("");  // Clear the selected file type
      e.target.reset();  
    } catch (error) {
      console.error("Error uploading updated Excel file:", error);
    }
  };

  const renderFileFields = () => {
    const sections = {
      coverFile: [
        "Book Cover PDF",
        "Book Cover PSD",
        "Cover Front Thumbnail",
        "Cover Back Thumbnail",
        "Book Cover Inside PDF (if applicable)",
        "Book Cover Inside PSD (if applicable)",
        "Cover Mockups",
        "Mockup 1",
        "Mockup 2",
        "Mockup 3",
        "Cover for GDP - PDF",
        "Cover for GDP - PSD",
      ],
      interiorFile: [
        "Upload PDF File",
        "Upload Open File (Indesign/Word/PageMaker/.pages)",
        "Interior File for GDP",
      ],
      epubFile: ["Upload EPUB file", "Upload KPF file"],
      posterFile: ["Upload the poster PDF"],
      globalFile: [
        "Cover for GDP - PDF",
        "Cover for GDP - PSD",
        "Revised Mockup",
        "Interior File for GDP",
      ],
    };

    return (
      <div className="space-y-4">
        {sections[fileType]?.map((field, index) => (
          <label key={index} className="block">
            <span className="text-gray-700">{field}:</span>
            <input type="file" name={field} onChange={handleChange2} className="block w-full border border-gray-300 rounded-md p-2 mt-1" />
          </label>
        ))}
      </div>
    );
  };

  return (
    <div className="center-box">
      <div className="max-w-xl mx-auto bg-white p-6 shadow-lg rounded-lg">
        <h2 className="text-xl font-semibold text-gray-800 mb-4">Final Files Submission</h2>
        <form onSubmit={handleSubmit} className="space-y-4">
          <label className="block">
            <span className="text-gray-700">File Type:</span>
            <select name="fileType" onChange={(e) => setFileType(e.target.value)} className="value">
              <option value="">-- Select --</option>
              <option value="coverFile">Cover File</option>
              <option value="interiorFile">Interior File</option>
              <option value="epubFile">ePub</option>
              <option value="posterFile">Poster</option>
              <option value="globalFile">Only Global - Cover and Interior</option>
            </select>
          </label>

          {fileType && (
            <div className="space-y-4 border-t pt-4">
              <label className="block">
                <span className="text-gray-700">File Status:</span>
                <select name="fileStatus" onChange={handleChange2} className="value">
                  <option value="new">New</option>
                  <option value="updated">Updated Version</option>
                </select>
              </label>

              <label className="block">
                <span className="text-gray-700">Book Title_Author Name_ISBN:</span>
                <input type="text" name="bookDetails" onChange={handleChange2} className="value" />
              </label>

              <label className="block">
                <span className="text-gray-700">BIN:</span>
                <input type="text" name="bin" onChange={handleChange2} className="value" />
              </label>

              {renderFileFields()}
            </div>
          )}

          <button type="submit" className="bg-blue-500 text-white px-4 py-2 rounded-md hover:bg-blue-600">
            Submit
          </button>
        </form>
      </div>
    </div>
  );
};

export default UploadForm;