/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import ReactDOM from 'react-dom/client';
import { GoogleGenAI, Type } from '@google/genai';
import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';

// Set up the PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://aistudiocdn.com/pdfjs-dist@^4.5.136/build/pdf.worker.min.mjs`;

// Helper to convert camelCase to Title Case for display
const toTitleCase = (str: string) => {
  return str.replace(/([A-Z])/g, ' $1').replace(/^./, (s) => s.toUpperCase());
};

// FIX: Add explicit return type React.ReactElement to avoid JSX namespace issues.
const App = (): React.ReactElement => {
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [abstractData, setAbstractData] = useState<any | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const addressSchema = {
    type: Type.OBJECT,
    properties: {
      name: { type: Type.STRING },
      atten: { type: Type.STRING },
      address: { type: Type.STRING },
      city: { type: Type.STRING },
      state: { type: Type.STRING },
      zipCode: { type: Type.STRING },
      email: { type: Type.STRING },
    },
  };

  const leaseAbstractSchema = {
    type: Type.OBJECT,
    description: "Detailed lease abstract data.",
    properties: {
      generalInformation: {
        type: Type.OBJECT,
        description: "Basic information about the lease agreement.",
        properties: {
          tenantName: { type: Type.STRING },
          dba: { type: Type.STRING, description: "Doing Business As" },
          suiteNumber: { type: Type.STRING },
          buildingNumber: { type: Type.STRING },
          storeNumber: { type: Type.STRING },
          shoppingCenterName: { type: Type.STRING },
          shoppingCenterAddress: { type: Type.STRING },
          premisesGLA: { type: Type.STRING, description: "Gross Leasable Area, numeric with commas." },
          guarantors: { type: Type.STRING },
          areSpousesGuarantors: { type: Type.STRING, description: "Options: Yes, No, Unsure" },
          isGuaranteeSeparate: { type: Type.STRING, description: "Is the guarantee signed separately from the lease? Options: Yes, No" },
        },
      },
      leaseAmendmentsReviewed: {
        type: Type.ARRAY,
        description: "A list of all lease documents and amendments reviewed.",
        items: {
          type: Type.OBJECT,
          properties: {
            dateOfDocument: { type: Type.STRING, description: "Format as MM/DD/YYYY" },
            documentReviewed: { type: Type.STRING },
            leaseSectionAndPage: { type: Type.STRING },
            issues: { type: Type.STRING },
          },
        },
      },
      leaseTermAndDates: {
        type: Type.OBJECT,
        description: "Key dates and term length for the lease.",
        properties: {
          leaseTerm: { type: Type.STRING, description: "e.g., '10 years and 3 months'" },
          leaseSectionAndPage: { type: Type.STRING },
          leaseExecutionDate: { type: Type.STRING, description: "Format as MM/DD/YYYY" },
          deliveryDate: { type: Type.STRING, description: "Format as MM/DD/YYYY" },
          leaseCommencementDate: { type: Type.STRING, description: "Format as MM/DD/YYYY" },
          openDate: { type: Type.STRING, description: "Format as MM/DD/YYYY" },
          rentCommencementDate: { type: Type.STRING, description: "Format as MM/DD/YYYY" },
          leaseExpirationDate: { type: Type.STRING, description: "Format as MM/DD/YYYY" },
        },
      },
      noticeAddresses: {
        type: Type.OBJECT,
        description: "Contact information for official notices.",
        properties: {
          leaseSectionAndPage: { type: Type.STRING },
          tenant: addressSchema,
          tenantsLawyer: addressSchema,
          franchisor: {
            type: Type.OBJECT,
            properties: {
              ...addressSchema.properties,
              leaseAddressType: { type: Type.STRING },
            },
          },
        },
      },
      billingAndCharges: {
        type: Type.OBJECT,
        description: "Financial details including rent and other charges.",
        properties: {
          leaseSectionAndPage: { type: Type.STRING },
          baseRentSchedule: {
            type: Type.ARRAY,
            description: "Schedule of base rent payments.",
            items: {
              type: Type.OBJECT,
              properties: {
                incomeCategory: { type: Type.STRING, description: "e.g., RNT" },
                effectiveDate: { type: Type.STRING, description: "Format as MM/DD/YYYY" },
                endDate: { type: Type.STRING, description: "Format as MM/DD/YYYY" },
                annualAmountPerSf: { type: Type.STRING, description: "Format as $XXX,XXX.XX" },
                annualTotal: { type: Type.STRING, description: "Format as $XXX,XXX.XX" },
                monthlyAmount: { type: Type.STRING, description: "Format as $XXX,XXX.XX" },
              },
            },
          },
          camTaxInsuranceFirstYear: { type: Type.STRING, description: "Amount to bill tenant for the first year CTI." },
          percentageRent: {
            type: Type.OBJECT,
            properties: {
              leaseSectionAndPage: { type: Type.STRING },
              reportingFrequency: { type: Type.STRING },
              naturalBreakpoint: { type: Type.STRING },
              unnaturalBreakpoint: { type: Type.STRING },
              salesYearEnd: { type: Type.STRING, description: "Format as DD/MM" },
              billingFrequency: { type: Type.STRING },
            },
          },
        },
      },
      leaseNotes: {
        type: Type.ARRAY,
        description: "Specific notes on lease clauses.",
        items: {
          type: Type.OBJECT,
          properties: {
            leaseSectionAndPage: { type: Type.STRING },
            reference: { type: Type.STRING, description: "e.g., Security Deposit, Use, Exclusive Use" },
            notes: { type: Type.STRING },
          },
        },
      },
      keyClauses: {
        type: Type.OBJECT,
        description: "Details on important lease clauses and covenants.",
        properties: {
          cotenancyRequirements: { type: Type.STRING },
          landlordKickout: { type: Type.STRING },
          tenantKickout: { type: Type.STRING },
          tenantGoDark: { type: Type.STRING },
          landlordRestrictions: { type: Type.STRING, description: "Leasing/No-Build restrictions" },
          assignmentAndSubletting: { type: Type.STRING },
          shoppingCenterAlterations: { type: Type.STRING },
          operatingCovenant: { type: Type.STRING, description: "Other than hours" },
          lateChargesNSFFee: { type: Type.STRING },
          defaultClause: { type: Type.STRING },
          guaranty: { type: Type.STRING },
          purchaseOptionROFR: { type: Type.STRING, description: "Purchase Option/Right of First Refusal/Right of First Offer" },
          marketingOrPromotionalFee: { type: Type.STRING },
          holdoverTerms: { type: Type.STRING },
          signage: { type: Type.STRING },
          estoppel: { type: Type.STRING },
          eminentDomainAndSubordination: { type: Type.STRING },
          damageOrDestruction: { type: Type.STRING },
          relocationRight: { type: Type.STRING },
        },
      },
      maintenanceAndReimbursement: {
        type: Type.OBJECT,
        description: "Responsibilities for maintenance, repairs, and reimbursements.",
        properties: {
          hvac: { type: Type.STRING },
          tenantAllowance: { type: Type.STRING },
          cam: {
            type: Type.OBJECT,
            description: "Common Area Maintenance details.",
            properties: {
              leaseSectionAndPage: { type: Type.STRING },
              prorataSharePercent: { type: Type.STRING },
              exclusions: { type: Type.STRING },
              paymentTerms: { type: Type.STRING },
              capitalRepairs: { type: Type.STRING },
              auditRight: { type: Type.STRING },
              adminFeeAllowedInCAM: { type: Type.STRING },
              propertyManagementFeeAllowedInCAM: { type: Type.STRING },
            },
          },
          realEstateTaxes: {
            type: Type.OBJECT,
            properties: {
              leaseSectionAndPage: { type: Type.STRING },
              prorataSharePercent: { type: Type.STRING },
              paymentTerms: { type: Type.STRING },
              appealRight: { type: Type.STRING },
              tenantPaysForAssessments: { type: Type.STRING },
            },
          },
          insurance: {
            type: Type.OBJECT,
            properties: {
              leaseSectionAndPage: { type: Type.STRING },
              insuranceReimbursed: { type: Type.STRING },
              tenantPaysDeductible: { type: Type.STRING },
              prorataSharePercent: { type: Type.STRING },
              rightToSelfInsure: { type: Type.STRING },
            },
          },
        },
      },
      tenantInsuranceInformation: {
        type: Type.OBJECT,
        description: "Tenant's insurance requirements.",
        properties: {
          leaseSectionAndPage: { type: Type.STRING },
          coverages: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                type: { type: Type.STRING, description: "e.g., Liability, Property Damage" },
                coverage: { type: Type.STRING },
              },
            },
          },
          deductibleRequirementInLease: { type: Type.STRING, description: "Yes/No" },
          comments: { type: Type.STRING },
        },
      },
    },
  };
  
  const handleFileChange = (files: FileList | null) => {
    if (!files) return;

    const newFiles = Array.from(files).filter(file => 
      file.type === 'application/pdf' && !selectedFiles.some(f => f.name === file.name)
    );
    
    if (newFiles.length > 0) {
      setSelectedFiles(prevFiles => [...prevFiles, ...newFiles].sort((a, b) => a.name.localeCompare(b.name)));
      setError(null);
    }
    
    const nonPdfFiles = Array.from(files).some(file => file.type !== 'application/pdf');
    if (nonPdfFiles) {
      setError('Only PDF files are accepted. Non-PDF files have been ignored.');
    }
  };

  const handleRemoveFile = (indexToRemove: number) => {
    const newFiles = selectedFiles.filter((_, index) => index !== indexToRemove);
    setSelectedFiles(newFiles);
    if (newFiles.length === 0) {
      setAbstractData(null); // Clear data if no files are left
    }
  };
  
  const extractTextFromPdf = async (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (event) => {
        if (!event.target?.result) {
          return reject(new Error("Failed to read file."));
        }
        try {
          const typedArray = new Uint8Array(event.target.result as ArrayBuffer);
          const pdf = await pdfjsLib.getDocument(typedArray).promise;
          let fullText = '';
          for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => ('str' in item ? item.str : '')).join(' ');
            fullText += pageText + '\n\n';
          }
          resolve(fullText);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = (error) => reject(error);
      reader.readAsArrayBuffer(file);
    });
  };

  const handleGenerateAbstract = async () => {
    if (selectedFiles.length === 0) {
      setError('Please upload one or more lease documents.');
      return;
    }
    setIsLoading(true);
    setError(null);
    setAbstractData(null);

    try {
      const leaseTexts = await Promise.all(
        selectedFiles.map(file => extractTextFromPdf(file))
      );
      const leaseText = leaseTexts.join('\n\n--- END OF DOCUMENT ---\n\n');

      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY! });
      const systemInstruction = `You are a world-class legal AI specializing in commercial lease abstraction. Your primary objective is to perform a meticulous and comprehensive review of all provided lease documents, which are concatenated and separated by '--- END OF DOCUMENT ---'. You must extract highly detailed information and populate the provided JSON schema with the utmost accuracy.

**Core Directives (Must be followed without exception):**

1.  **Chronological Supremacy & Conflict Resolution:**
    *   The documents are provided in chronological order. You MUST analyze them sequentially to build a timeline of the lease.
    *   Later documents (e.g., amendments, addendums) supersede earlier ones. If a clause is amended, you MUST extract the information from the **latest effective amendment**.
    *   The final abstract MUST reflect the current, legally binding state of the lease. Do not include outdated or superseded information. For example, if Base Rent is changed in the Second Amendment, the 'Base Rent Schedule' must only show the new, current schedule from that amendment.

2.  **Extreme Detail & Verbatim Extraction:**
    *   Your goal is comprehensiveness, not brevity. **Do not summarize.**
    *   For all clauses, covenants, restrictions, and significant terms (e.g., Use, Exclusivity, Cotenancy, CAM Exclusions), you MUST extract the **full, verbatim text** from the source document. Short, one-sentence summaries are unacceptable.
    *   Provide all available information for every field. If a detail seems minor, include it. The user requires a complete picture.

3.  **Meticulous Sourcing and Citation:**
    *   For every piece of information extracted, you MUST cite its source, including the document name (e.g., "Original Lease," "Second Amendment"), the section number, and the page number. This is non-negotiable. Example: "Second Amendment, Section 3.a, Page 2".

4.  **Absolute Schema Adherence:**
    *   Strictly follow the provided JSON schema, including all data types and formatting rules (Dates: MM/DD/YYYY; Currency: $XXX,XXX.XX; Square Footage: Numeric with commas).

5.  **"Not Provided" as a Last Resort:**
    *   Only use "Not Provided" after you have exhaustively searched all documents and are certain the information does not exist. Before concluding information is missing, double-check all amendments and exhibits.`;


      const response = await ai.models.generateContent({
        model: 'gemini-2.5-flash',
        contents: leaseText,
        config: {
          systemInstruction: systemInstruction,
          responseMimeType: 'application/json',
          responseSchema: leaseAbstractSchema,
        },
      });

      const jsonResponse = JSON.parse(response.text);
      setAbstractData(jsonResponse);
    } catch (e: any) {
      console.error(e);
      setError(`An error occurred: ${e.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  const handleExportToExcel = () => {
    if (!abstractData) return;

    const flattenDataForExcel = (data: any) => {
      const rows: (string | number)[][] = [];
      const process = (obj: any, prefix = '') => {
        if (typeof obj !== 'object' || obj === null) return;

        Object.entries(obj).forEach(([key, value]) => {
          const newPrefix = prefix ? `${prefix} > ${toTitleCase(key)}` : toTitleCase(key);
          if (typeof value === 'object' && value !== null) {
            if (Array.isArray(value)) {
              rows.push([newPrefix]);
              value.forEach((item, index) => {
                rows.push([`${newPrefix} [${index + 1}]`]);
                process(item, `  `); // Indent array items
              });
            } else {
              process(value, newPrefix);
            }
          } else {
            // FIX: Explicitly convert the value to a string to prevent type errors when pushing to the `rows` array, which expects `string | number`.
            rows.push([newPrefix, String(value ?? '')]);
          }
        });
      };
      process(data);
      return rows;
    };

    const data = flattenDataForExcel(abstractData);
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 45 }, { wch: 60 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Lease Abstract');
    XLSX.writeFile(wb, 'Lease_Abstract.xlsx');
  };

  // Drag and drop handlers
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>) => { e.preventDefault(); setIsDragging(false); };
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileChange(e.dataTransfer.files);
      e.dataTransfer.clearData();
    }
  };

  // FIX: Change parameter type to 'any' to handle various data types from the API response and prevent type errors.
  const renderValueHighlight = (value: any) => {
    // FIX: Safely convert value to string, handling null/undefined.
    const stringValue = String(value ?? '');
    if (stringValue.toLowerCase().includes('not provided') || stringValue.toLowerCase().includes('conflicting info')) {
      return <span className="highlight">{stringValue}</span>;
    }
    return stringValue;
  };
  
  // FIX: Change return type from JSX.Element to React.ReactElement to avoid JSX namespace issues.
  const renderAbstractData = (data: any): React.ReactElement => {
    if (typeof data !== 'object' || data === null) {
      return <div className="value-wrapper">{renderValueHighlight(data)}</div>;
    }
  
    if (Array.isArray(data)) {
      return (
        <div className="result-array">
          {data.map((item, index) => (
            <div key={index} className="result-card nested-card">
              <h4>Item {index + 1}</h4>
              {renderAbstractData(item)}
            </div>
          ))}
        </div>
      );
    }
  
    return (
      <ul className="result-list">
        {Object.entries(data).map(([key, value]) => (
          <li key={key}>
            <strong>{toTitleCase(key)}:</strong>
            {renderAbstractData(value)}
          </li>
        ))}
      </ul>
    );
  };
  

  return (
    <div className="app-container">
      <header>
        <h1>Lease Abstractor</h1>
        <p>AI-Powered Lease Data Extraction</p>
      </header>
      <main className="main-content">
        <div className="input-panel">
          <h2>Lease Documents</h2>
          <div 
            className={`file-dropzone ${isDragging ? 'drag-over' : ''}`}
            onClick={() => fileInputRef.current?.click()}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={(e) => handleFileChange(e.target.files)}
              accept="application/pdf"
              style={{ display: 'none' }}
              disabled={isLoading}
              multiple
            />
            {selectedFiles.length > 0 ? (
              <div className="file-list-container">
                <ul className="file-list">
                  {selectedFiles.map((file, index) => (
                    <li key={file.name} className="file-item">
                      <i className="fa-solid fa-file-pdf"></i>
                      <span>{file.name}</span>
                      <button 
                        className="remove-file-btn"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleRemoveFile(index);
                        }}
                        aria-label={`Remove ${file.name}`}
                      >
                        &times;
                      </button>
                    </li>
                  ))}
                </ul>
                <button
                  className="clear-all-btn"
                  onClick={(e) => {
                    e.stopPropagation();
                    setSelectedFiles([]);
                    setAbstractData(null);
                    if (fileInputRef.current) fileInputRef.current.value = '';
                  }}
                  aria-label="Remove all files"
                >
                  Clear All
                </button>
              </div>
            ) : (
              <div className="drop-prompt">
                 <i className="fa-solid fa-cloud-arrow-up"></i>
                 <p>Drag & drop your PDFs here, or <strong>click to browse</strong>.</p>
              </div>
            )}
          </div>
          <button onClick={handleGenerateAbstract} disabled={selectedFiles.length === 0 || isLoading} aria-live="polite">
            {isLoading ? (
              <>
                <div className="spinner" aria-hidden="true"></div>
                Generating...
              </>
            ) : (
              'Generate Abstract'
            )}
          </button>
        </div>
        <div className="output-panel">
          <div className="output-header">
            <h2>Lease Abstract</h2>
            {abstractData && !isLoading && (
              <button className="export-button" onClick={handleExportToExcel}>
                <i className="fa-solid fa-file-excel"></i> Export to Excel
              </button>
            )}
          </div>
          <div className="results-container" aria-live="polite">
            {isLoading && (
               <div className="skeleton-loader">
                 <div className="skeleton-card"></div>
                 <div className="skeleton-card"></div>
                 <div className="skeleton-card"></div>
               </div>
            )}
            {error && <div className="error-message" role="alert">{error}</div>}
            {abstractData ? (
              Object.entries(abstractData).map(([sectionTitle, sectionData]) => (
                <div key={sectionTitle} className="result-card">
                  <h3>{toTitleCase(sectionTitle)}</h3>
                  {renderAbstractData(sectionData)}
                </div>
              ))
            ) : (
              !isLoading && !error && <p className="placeholder-text">Your extracted lease details will appear here.</p>
            )}
          </div>
        </div>
      </main>
    </div>
  );
};

const root = ReactDOM.createRoot(document.getElementById('root')!);
root.render(<App />);