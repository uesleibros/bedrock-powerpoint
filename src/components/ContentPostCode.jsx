"use client";

import React, { useEffect } from "react";
import hljs from "highlight.js";
import "highlight.js/styles/github.css";

export default function ContentPostCode({ htmlContent }) {
  useEffect(() => {
    hljs.highlightAll();
  }, [htmlContent]);

  return (<div className="prose max-w-[800px]" dangerouslySetInnerHTML={{ __html: htmlContent }} />);
}