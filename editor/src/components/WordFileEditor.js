import './WordFileEditor.css';
import "react-draft-wysiwyg/dist/react-draft-wysiwyg.css";
import React, { useEffect, useState } from 'react';
import { Editor } from 'react-draft-wysiwyg';
import { EditorState, ContentState, convertFromHTML, convertToRaw } from 'draft-js';
import 'react-draft-wysiwyg/dist/react-draft-wysiwyg.css';
import axios from 'axios';

function WordFileEditor() {
  const [editorState, setEditorState] = useState(() => EditorState.createEmpty());
  const [htmlContent, setHtmlContent] = useState();
  
  const parseDocxToHtml = async (file) => {
    var form = new FormData();
    form.append("file", file, "file-sample_100kB.docx");
    try {
      const response = await axios.post('http://127.0.0.1:8000/parse-docx-to-html', form);
      setHtmlContent(response.data.html_content);
    } catch (error) {
      console.error('Error uploading DOCX file:', error);
    }
  }
  useEffect(() => {
    if (htmlContent) {
      const blocksFromHTML = convertFromHTML(htmlContent);
      const contentState = ContentState.createFromBlockArray(
        blocksFromHTML.contentBlocks,
        blocksFromHTML.entityMap
      );
      setEditorState(EditorState.createWithContent(contentState));
    }
  }, [htmlContent]);

  const handleFileUpload = () => {
    document.getElementById("UploadFile").click();
  };


  const downloadFile = async () => {
    const content = convertToRaw(editorState.getCurrentContent());
    try {
      const response = await axios.post('http://127.0.0.1:8000/parse-to-docx', {
        content: content
      }
        , {
          responseType: 'blob',
        });

      // Create a blob URL and trigger the download
      const uploadedFileName = document.getElementById('UploadFile').files[0].name;
      const blob = new Blob([response.data], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = uploadedFileName;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Error downloading DOCX file', error);
    }
  }

  // Handle file input change
  const handleFileChange = event => {
    const selectedFile = event.target.files[0];
    if (selectedFile) {
      parseDocxToHtml(selectedFile);
    }
  };

  // Handle changes in the editor
  const handleEditorChange = newEditorState => {
    setEditorState(newEditorState);
    // You can save the updated content to a file or perform other actions as needed.
  };
  return (
    <div>
      <div className='mainInterface'>
        <Editor
          className="wholeEditor"
          toolbarClassName="toolbarClassName"
          wrapperClassName="wrapperClassName"
          editorClassName="editorClassName"
          editorState={editorState}
          onEditorStateChange={handleEditorChange}
        />
      </div>
      <hr />

      <input type="file" id='UploadFile' onChange={handleFileChange} hidden />
      <button onClick={handleFileUpload}>Upload Docx File</button>
      <button onClick={downloadFile}>Download Docx File</button>
    </div>
  );
}

export default WordFileEditor;
