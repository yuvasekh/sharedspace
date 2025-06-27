"use client"

import { useState, useRef, useCallback, useEffect } from "react"
import * as XLSX from "xlsx"
import {
  Upload,
  Play,
  CheckCircle,
  AlertCircle,
  FileText,
  Server,
  Database,
  Activity,
  Download,
  RefreshCw,
  Eye,
  X,
  Building2,
  Shield,
  Clock,
  Users,
  TrendingUp,
  Settings,
  HelpCircle,
  Wifi,
  WifiOff,
  ExternalLink,
  Square,
} from "lucide-react"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Textarea } from "@/components/ui/textarea"
import { Badge } from "@/components/ui/badge"
import { Progress } from "@/components/ui/progress"
import { Alert, AlertDescription } from "@/components/ui/alert"
import { Separator } from "@/components/ui/separator"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"

const Unit4XlDashboard = () => {
  const [instanceUrl, setInstanceUrl] = useState("https://demo-spaces.gainsightcloud.com")
  const [cookie, setCookie] = useState("")
  const [xlFile, setXlFile] = useState(null)
  const [xlPreview, setXlPreview] = useState([])
  const [worksheetNames, setWorksheetNames] = useState([])
  const [selectedWorksheet, setSelectedWorksheet] = useState("")
  const [isProcessing, setIsProcessing] = useState(false)
  const [isParsingFile, setIsParsingFile] = useState(false)
  const [apiResponse, setApiResponse] = useState(null)
  const [error, setError] = useState(null)
  const [showCookieModal, setShowCookieModal] = useState(false)
  const [connectionStatus, setConnectionStatus] = useState("online")
  const [hyperlinkCount, setHyperlinkCount] = useState(0)

  // New states for long-running process handling
  const [processingProgress, setProcessingProgress] = useState(0)
  const [processingStatus, setProcessingStatus] = useState("")
  const [estimatedTimeRemaining, setEstimatedTimeRemaining] = useState(null)
  const [processStartTime, setProcessStartTime] = useState(null)
  const [jobId, setJobId] = useState(null)
  const [isPolling, setIsPolling] = useState(false)
  const [canCancel, setCanCancel] = useState(false)

  const fileInputRef = useRef(null)
  const abortControllerRef = useRef(null)
  const pollingIntervalRef = useRef(null)

  // File size limit (10MB to match backend)
  const MAX_FILE_SIZE = 10 * 1024 * 1024
  // Extended timeout for long-running processes (30 minutes)
  const EXTENDED_TIMEOUT = 30 * 60 * 1000

  // Enhanced hyperlink extraction function
  const extractHyperlinksFromWorksheet = (worksheet, jsonData) => {
    const processedData = []
    let totalHyperlinks = 0

    jsonData.forEach((row, rowIndex) => {
      const processedRow = { ...row }

      Object.keys(processedRow).forEach((columnName, colIndex) => {
        const cellAddress = XLSX.utils.encode_cell({
          c: colIndex,
          r: rowIndex + 1,
        })

        const cell = worksheet[cellAddress]

        if (cell) {
          let hasHyperlink = false
          let linkUrl = null

          if (cell.l && cell.l.Target) {
            linkUrl = cell.l.Target
            hasHyperlink = true
          } else if (cell.l && cell.l.Hyperlink) {
            linkUrl = cell.l.Hyperlink
            hasHyperlink = true
          } else if (typeof cell.v === "string" && (cell.v.startsWith("http://") || cell.v.startsWith("https://"))) {
            linkUrl = cell.v
            hasHyperlink = true
          } else if (cell.f && cell.f.includes("HYPERLINK")) {
            const hyperlinkMatch = cell.f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i)
            if (hyperlinkMatch) {
              linkUrl = hyperlinkMatch[1]
              hasHyperlink = true
            }
          }

          if (hasHyperlink && linkUrl) {
            processedRow[columnName] = {
              text: processedRow[columnName] || cell.v || "",
              link: linkUrl,
              hasHyperlink: true,
            }
            totalHyperlinks++
          }
        }
      })

      processedData.push(processedRow)
    })

    setHyperlinkCount(totalHyperlinks)
    return processedData
  }

  // Polling function to check job status
  const pollJobStatus = useCallback(async (jobId) => {
    try {
      const response = await fetch(`https://sharedspace-w4ka.onrender.com/api/job-status/${jobId}`, {
        method: "GET",
        headers: {
          "Content-Type": "application/json",
        },
      })

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`)
      }

      const data = await response.json()

      // Update progress and status
      setProcessingProgress(data.progress || 0)
      setProcessingStatus(data.status || "Processing...")

      if (data.estimatedTimeRemaining) {
        setEstimatedTimeRemaining(data.estimatedTimeRemaining)
      }

      // Check if job is complete
      if (data.completed) {
        setIsPolling(false)
        setIsProcessing(false)
        setCanCancel(false)

        if (data.success) {
          setApiResponse(data.result)
          setProcessingStatus("Completed successfully!")
        } else {
          setError(data.error || "Processing failed")
          setProcessingStatus("Processing failed")
        }

        if (pollingIntervalRef.current) {
          clearInterval(pollingIntervalRef.current)
          pollingIntervalRef.current = null
        }
      }
    } catch (err) {
      console.error("Polling error:", err)
      // Continue polling on error, but log it
    }
  }, [])

  // Start polling for job status
  const startPolling = useCallback(
    (jobId) => {
      setIsPolling(true)
      setJobId(jobId)

      // Poll every 5 seconds
      pollingIntervalRef.current = setInterval(() => {
        pollJobStatus(jobId)
      }, 5000)

      // Initial poll
      pollJobStatus(jobId)
    },
    [pollJobStatus],
  )

  // Cancel job function
  const cancelJob = useCallback(async () => {
    if (!jobId) return

    try {
      const response = await fetch(`https://sharedspace-w4ka.onrender.com/api/cancel-job/${jobId}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
      })

      if (response.ok) {
        setIsProcessing(false)
        setIsPolling(false)
        setCanCancel(false)
        setProcessingStatus("Cancelled by user")

        if (pollingIntervalRef.current) {
          clearInterval(pollingIntervalRef.current)
          pollingIntervalRef.current = null
        }
      }
    } catch (err) {
      console.error("Cancel job error:", err)
    }
  }, [jobId])

  // Cleanup polling on unmount
  useEffect(() => {
    return () => {
      if (pollingIntervalRef.current) {
        clearInterval(pollingIntervalRef.current)
      }
      if (abortControllerRef.current) {
        abortControllerRef.current.abort()
      }
    }
  }, [])

  // Check if URL is valid
  const isValidUrl = (string) => {
    try {
      new URL(string)
      return true
    } catch (_) {
      return false
    }
  }

  // Check if file is Excel format
  const isExcelFile = (file) => {
    const validTypes = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
      "application/vnd.ms-excel.sheet.macroEnabled.12",
    ]
    const validExtensions = [".xlsx", ".xls", ".xlsm"]
    const fileExtension = file.name.toLowerCase().substring(file.name.lastIndexOf("."))

    return validTypes.includes(file.type) || validExtensions.includes(fileExtension)
  }

  // Enhanced error handling
  const handleError = useCallback((error, context = "") => {
    console.error(`Error in ${context}:`, error)

    let errorMessage = "An unexpected error occurred"

    if (error.name === "NetworkError" || error.message.includes("fetch")) {
      errorMessage = "Network connection failed. Please check your internet connection and try again."
      setConnectionStatus("offline")
    } else if (error.message.includes("timeout")) {
      errorMessage = "Request timed out. The server may be busy, please try again."
    } else if (error.message.includes("413")) {
      errorMessage = "File too large. Please select a file smaller than 10MB."
    } else if (error.message.includes("400")) {
      errorMessage = "Invalid request. Please check your input data and try again."
    } else if (error.message.includes("401") || error.message.includes("403")) {
      errorMessage = "Authentication failed. Please check your cookie and try again."
    } else if (error.message.includes("404")) {
      errorMessage = "Server endpoint not found. Please check the instance URL."
    } else if (error.message.includes("500")) {
      errorMessage = "Server error occurred. Please try again later or contact support."
    } else if (error.message.includes("parsing")) {
      errorMessage = "Error parsing Excel file. Please ensure the file is not corrupted and try again."
    } else if (error.message) {
      errorMessage = error.message
    }

    setError(errorMessage)
  }, [])

  // File upload handler
  const handleFileUpload = useCallback(
    (event) => {
      const file = event.target.files?.[0]

      if (!file) {
        setError("No file selected")
        return
      }

      setError(null)
      setXlPreview([])
      setWorksheetNames([])
      setSelectedWorksheet("")
      setHyperlinkCount(0)

      if (!isExcelFile(file)) {
        setError("Please select a valid Excel file (.xlsx, .xls, .xlsm)")
        return
      }

      if (file.size > MAX_FILE_SIZE) {
        setError(`File size (${(file.size / 1024 / 1024).toFixed(2)}MB) exceeds the maximum limit of 10MB`)
        return
      }

      setXlFile(file)
      setIsParsingFile(true)

      const reader = new FileReader()

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result)
          const workbook = XLSX.read(data, { type: "array" })

          if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
            throw new Error("No worksheets found in the Excel file")
          }

          const sheetNames = workbook.SheetNames
          setWorksheetNames(sheetNames)

          const firstSheetName = sheetNames[0]
          setSelectedWorksheet(firstSheetName)

          const worksheet = workbook.Sheets[firstSheetName]
          if (!worksheet) {
            throw new Error(`Worksheet "${firstSheetName}" not found`)
          }

          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

          if (jsonData.length === 0) {
            setError("The selected worksheet appears to be empty")
            return
          }

          if (jsonData.length === 1) {
            setError("The worksheet only contains headers. Please ensure there is data to process.")
            return
          }

          const headers = jsonData[0]
          if (!headers || headers.length === 0) {
            throw new Error("No headers found in the worksheet")
          }

          const dataRows = jsonData.slice(1, 6).map((row, rowIndex) => {
            const obj = {}
            headers.forEach((header, index) => {
              obj[header] = row[index] || ""
            })
            return obj
          })

          const processedPreview = extractHyperlinksFromWorksheet(worksheet, dataRows)
          setXlPreview(processedPreview)
        } catch (err) {
          handleError(err, "file parsing")
        } finally {
          setIsParsingFile(false)
        }
      }

      reader.onerror = () => {
        setError("Failed to read the file. The file may be corrupted.")
        setIsParsingFile(false)
      }

      reader.readAsArrayBuffer(file)
    },
    [handleError],
  )

  // Worksheet change handler
  const handleWorksheetChange = useCallback(
    (sheetName) => {
      if (!xlFile) return

      setSelectedWorksheet(sheetName)
      setIsParsingFile(true)
      setError(null)
      setHyperlinkCount(0)

      const reader = new FileReader()
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result)
          const workbook = XLSX.read(data, { type: "array" })
          const worksheet = workbook.Sheets[sheetName]

          if (!worksheet) {
            throw new Error(`Worksheet "${sheetName}" not found`)
          }

          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

          if (jsonData.length === 0) {
            setError(`Worksheet "${sheetName}" appears to be empty`)
            return
          }

          if (jsonData.length === 1) {
            setError(`Worksheet "${sheetName}" only contains headers. Please ensure there is data to process.`)
            return
          }

          const headers = jsonData[0]
          const dataRows = jsonData.slice(1, 6).map((row) => {
            const obj = {}
            headers.forEach((header, index) => {
              obj[header] = row[index] || ""
            })
            return obj
          })

          const processedPreview = extractHyperlinksFromWorksheet(worksheet, dataRows)
          setXlPreview(processedPreview)
        } catch (err) {
          handleError(err, "worksheet parsing")
        } finally {
          setIsParsingFile(false)
        }
      }

      reader.onerror = () => {
        setError("Failed to read the worksheet. Please try again.")
        setIsParsingFile(false)
      }

      reader.readAsArrayBuffer(xlFile)
    },
    [xlFile, handleError],
  )

  // Form validation
  const validateForm = useCallback(() => {
    if (!instanceUrl.trim()) {
      setError("Instance URL is required")
      return false
    }

    if (!isValidUrl(instanceUrl)) {
      setError("Please enter a valid URL (e.g., https://your-instance.example.com)")
      return false
    }

    if (!cookie.trim()) {
      setError("Authentication cookie is required")
      return false
    }

    if (!xlFile) {
      setError("Please select an Excel file to upload")
      return false
    }

    return true
  }, [instanceUrl, cookie, xlFile])

  // Enhanced submit handler with long-running process support
  const handleSubmit = useCallback(async () => {
    if (!validateForm()) {
      return
    }

    setIsProcessing(true)
    setError(null)
    setApiResponse(null)
    setConnectionStatus("online")
    setProcessingProgress(0)
    setProcessingStatus("Initializing...")
    setProcessStartTime(Date.now())
    setCanCancel(true)

    // Create abort controller with extended timeout
    abortControllerRef.current = new AbortController()
    const timeoutId = setTimeout(() => {
      if (abortControllerRef.current) {
        abortControllerRef.current.abort()
      }
    }, EXTENDED_TIMEOUT)

    try {
      const formData = new FormData()
      formData.append("file", xlFile)
      formData.append("url", instanceUrl.trim())
      formData.append("cookie", cookie.trim())
      formData.append("async", "true") // Request async processing

      if (selectedWorksheet) {
        formData.append("worksheet", selectedWorksheet)
      }

      setProcessingStatus("Uploading file and starting processing...")

      const response = await fetch("http://localhost:3000/api/xlsxfileupload", {
        method: "POST",
        body: formData,
        signal: abortControllerRef.current.signal,
      })

      clearTimeout(timeoutId)

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}))
        throw new Error(errorData.message || `HTTP ${response.status}: ${response.statusText}`)
      }

      const data = await response.json()

      if (data.jobId) {
        // Async processing - start polling
        setProcessingStatus("Processing started. Monitoring progress...")
        startPolling(data.jobId)
      } else if (data.success) {
        // Immediate response (for smaller files)
        setApiResponse(data)
        setIsProcessing(false)
        setCanCancel(false)
        setProcessingStatus("Completed successfully!")
        setProcessingProgress(100)
      } else {
        throw new Error(data.message || "Processing failed")
      }
    } catch (err) {
      clearTimeout(timeoutId)
      setIsProcessing(false)
      setCanCancel(false)

      if (err.name === "AbortError") {
        setProcessingStatus("Request cancelled")
        handleError(new Error("Request was cancelled"), "submit")
      } else {
        handleError(err, "submit")
      }
    }
  }, [validateForm, xlFile, instanceUrl, cookie, selectedWorksheet, handleError, startPolling])

  // Reset function
  const resetForm = useCallback(() => {
    // Cancel any ongoing operations
    if (abortControllerRef.current) {
      abortControllerRef.current.abort()
    }
    if (pollingIntervalRef.current) {
      clearInterval(pollingIntervalRef.current)
      pollingIntervalRef.current = null
    }

    setInstanceUrl("https://demo-spaces.gainsightcloud.com")
    setCookie("")
    setXlFile(null)
    setXlPreview([])
    setWorksheetNames([])
    setSelectedWorksheet("")
    setApiResponse(null)
    setError(null)
    setConnectionStatus("online")
    setHyperlinkCount(0)
    setIsProcessing(false)
    setIsPolling(false)
    setProcessingProgress(0)
    setProcessingStatus("")
    setEstimatedTimeRemaining(null)
    setProcessStartTime(null)
    setJobId(null)
    setCanCancel(false)

    if (fileInputRef.current) {
      fileInputRef.current.value = ""
    }
  }, [])

  // Format time remaining
  const formatTimeRemaining = (seconds) => {
    if (!seconds) return null
    const minutes = Math.floor(seconds / 60)
    const remainingSeconds = seconds % 60
    return `${minutes}:${remainingSeconds.toString().padStart(2, "0")}`
  }

  // Calculate elapsed time
  const getElapsedTime = () => {
    if (!processStartTime) return null
    const elapsed = Math.floor((Date.now() - processStartTime) / 1000)
    return formatTimeRemaining(elapsed)
  }

  // Status color helper
  const getStatusColor = (status) => {
    switch (status) {
      case "Success":
        return "bg-emerald-50 text-emerald-700 border-emerald-200"
      case "Partial":
        return "bg-amber-50 text-amber-700 border-amber-200"
      case "Failed":
        return "bg-red-50 text-red-700 border-red-200"
      default:
        return "bg-slate-50 text-slate-700 border-slate-200"
    }
  }

  // Status icon helper
  const getStatusIcon = (status) => {
    switch (status) {
      case "Success":
        return <CheckCircle className="w-4 h-4" />
      case "Partial":
        return <AlertCircle className="w-4 h-4" />
      case "Failed":
        return <X className="w-4 h-4" />
      default:
        return <RefreshCw className="w-4 h-4" />
    }
  }

  // Helper function to render cell content with hyperlinks
  const renderCellContent = (cell) => {
    if (typeof cell === "object" && cell.hasHyperlink) {
      return (
        <div className="flex items-center space-x-1">
          <a
            href={cell.link}
            target="_blank"
            rel="noopener noreferrer"
            className="text-blue-600 hover:text-blue-800 underline flex items-center"
          >
            {cell.text.length > 25 ? `${cell.text.substring(0, 25)}...` : cell.text}
            <ExternalLink className="w-3 h-3 ml-1" />
          </a>
        </div>
      )
    }

    const cellText = String(cell)
    return cellText.length > 30 ? `${cellText.substring(0, 30)}...` : cellText
  }

  // Calculate statistics
  const stats = apiResponse?.processedData
    ? {
        total: apiResponse.processedData.length,
        success: apiResponse.processedData.filter((r) => r.status === "Success").length,
        partial: apiResponse.processedData.filter((r) => r.status === "Partial").length,
        failed: apiResponse.processedData.filter((r) => r.status === "Failed").length,
      }
    : { total: 0, success: 0, partial: 0, failed: 0 }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-white to-blue-50">
      {/* Header */}
      <div className="bg-white border-b border-slate-200 shadow-sm">
        <div className="max-w-7xl mx-auto px-6 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-4">
              <div className="flex items-center space-x-3">
                <div className="w-10 h-10 bg-gradient-to-r from-green-600 to-emerald-600 rounded-lg flex items-center justify-center">
                  <Building2 className="w-6 h-6 text-white" />
                </div>
                <div>
                  <h1 className="text-xl font-semibold text-slate-900">Unit4 Excel Processing</h1>
                  <p className="text-sm text-slate-500">Long-Running Process Support - Excel with Hyperlinks</p>
                </div>
              </div>
            </div>
            <div className="flex items-center space-x-3">
              <Badge variant="outline" className="bg-green-50 text-green-700 border-green-200">
                <Shield className="w-3 h-3 mr-1" />
                Secure Environment
              </Badge>
              {hyperlinkCount > 0 && (
                <Badge variant="outline" className="bg-blue-50 text-blue-700 border-blue-200">
                  <ExternalLink className="w-3 h-3 mr-1" />
                  {hyperlinkCount} Hyperlinks Found
                </Badge>
              )}
              <Badge
                variant="outline"
                className={
                  connectionStatus === "online"
                    ? "bg-green-50 text-green-700 border-green-200"
                    : "bg-red-50 text-red-700 border-red-200"
                }
              >
                {connectionStatus === "online" ? (
                  <Wifi className="w-3 h-3 mr-1" />
                ) : (
                  <WifiOff className="w-3 h-3 mr-1" />
                )}
                {connectionStatus === "online" ? "Online" : "Offline"}
              </Badge>
              <Button variant="ghost" size="sm">
                <HelpCircle className="w-4 h-4 mr-2" />
                Help
              </Button>
            </div>
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto px-6 py-8">
        {/* Error Alert */}
        {error && (
          <Alert className="mb-6 border-red-200 bg-red-50">
            <AlertCircle className="h-4 w-4 text-red-600" />
            <AlertDescription className="text-red-800">{error}</AlertDescription>
          </Alert>
        )}

        {/* Processing Status Alert */}
        {(isProcessing || isPolling) && (
          <Alert className="mb-6 border-blue-200 bg-blue-50">
            <Activity className="h-4 w-4 text-blue-600 animate-pulse" />
            <AlertDescription className="text-blue-800">
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <span>{processingStatus}</span>
                  {canCancel && (
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={cancelJob}
                      className="ml-4 h-7 px-3 text-xs bg-transparent"
                    >
                      <Square className="w-3 h-3 mr-1" />
                      Cancel
                    </Button>
                  )}
                </div>
                <div className="space-y-1">
                  <Progress value={processingProgress} className="h-2" />
                  <div className="flex justify-between text-xs text-blue-600">
                    <span>{processingProgress.toFixed(1)}% complete</span>
                    <div className="flex space-x-4">
                      {getElapsedTime() && <span>Elapsed: {getElapsedTime()}</span>}
                      {estimatedTimeRemaining && (
                        <span>Est. remaining: {formatTimeRemaining(estimatedTimeRemaining)}</span>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            </AlertDescription>
          </Alert>
        )}

        {/* Main Content */}
        <div className="grid lg:grid-cols-3 gap-8">
          {/* Configuration Panel */}
          <div className="lg:col-span-1">
            <Card className="shadow-lg border-0 bg-white">
              <CardHeader className="pb-4">
                <CardTitle className="flex items-center text-lg font-semibold">
                  <Settings className="w-5 h-5 mr-3 text-green-600" />
                  Configuration
                </CardTitle>
                <p className="text-sm text-slate-500">Set up your processing parameters</p>
              </CardHeader>
              <CardContent className="space-y-6">
                {/* Instance URL */}
                <div className="space-y-2">
                  <label className="text-sm font-medium text-slate-700 flex items-center">
                    <Server className="w-4 h-4 mr-2 text-slate-500" />
                    Instance URL *
                  </label>
                  <Input
                    type="url"
                    value={instanceUrl}
                    onChange={(e) => setInstanceUrl(e.target.value)}
                    placeholder="https://your-instance.example.com"
                    className="h-11 border-slate-200 focus:border-green-500 focus:ring-green-500"
                    disabled={isProcessing || isPolling}
                  />
                  {instanceUrl && !isValidUrl(instanceUrl) && (
                    <p className="text-xs text-red-600">Please enter a valid URL</p>
                  )}
                </div>

                {/* Cookie */}
                <div className="space-y-2">
                  <label className="text-sm font-medium text-slate-700 flex items-center">
                    <Shield className="w-4 h-4 mr-2 text-slate-500" />
                    Authentication Cookie *
                  </label>
                  <div className="relative">
                    <Textarea
                      value={cookie}
                      onChange={(e) => setCookie(e.target.value)}
                      placeholder="Paste your authentication cookie here..."
                      className="min-h-[80px] pr-12 border-slate-200 focus:border-green-500 focus:ring-green-500"
                      disabled={isProcessing || isPolling}
                    />
                    <Button
                      variant="ghost"
                      size="sm"
                      className="absolute top-2 right-2 h-8 w-8 p-0"
                      onClick={() => setShowCookieModal(true)}
                      disabled={isProcessing || isPolling}
                    >
                      <Eye className="w-4 h-4" />
                    </Button>
                  </div>
                </div>

                {/* File Upload */}
                <div className="space-y-2">
                  <label className="text-sm font-medium text-slate-700 flex items-center">
                    <FileText className="w-4 h-4 mr-2 text-slate-500" />
                    Excel File *
                  </label>
                  <div className="relative">
                    <input
                      ref={fileInputRef}
                      type="file"
                      accept=".xlsx,.xls,.xlsm"
                      onChange={handleFileUpload}
                      className="hidden"
                      disabled={isProcessing || isParsingFile || isPolling}
                    />
                    <div
                      onClick={() => !isProcessing && !isParsingFile && !isPolling && fileInputRef.current?.click()}
                      className={`w-full p-6 border-2 border-dashed rounded-lg transition-all duration-200 cursor-pointer group ${
                        isProcessing || isParsingFile || isPolling
                          ? "border-slate-200 bg-slate-50 cursor-not-allowed"
                          : "border-slate-300 hover:border-green-400 bg-slate-50 hover:bg-green-50"
                      }`}
                    >
                      <div className="text-center">
                        {isParsingFile ? (
                          <RefreshCw className="w-8 h-8 text-green-500 mx-auto mb-2 animate-spin" />
                        ) : (
                          <Upload
                            className={`w-8 h-8 mx-auto mb-2 transition-colors ${
                              isProcessing || isPolling ? "text-slate-300" : "text-slate-400 group-hover:text-green-500"
                            }`}
                          />
                        )}
                        <p className="text-slate-600 font-medium text-sm">
                          {isParsingFile
                            ? "Parsing Excel file..."
                            : xlFile
                              ? xlFile.name
                              : "Click to upload Excel file"}
                        </p>
                        <p className="text-xs text-slate-400 mt-1">
                          {xlFile
                            ? `${(xlFile.size / 1024).toFixed(2)} KB • ${worksheetNames.length} worksheet${worksheetNames.length !== 1 ? "s" : ""} • ${hyperlinkCount} hyperlinks`
                            : "Supports .xlsx, .xls, .xlsm files (max 10MB)"}
                        </p>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Worksheet Selection */}
                {worksheetNames.length > 1 && (
                  <div className="space-y-2">
                    <label className="text-sm font-medium text-slate-700">Select Worksheet</label>
                    <div className="grid grid-cols-1 gap-2 max-h-32 overflow-y-auto">
                      {worksheetNames.map((sheetName) => (
                        <Button
                          key={sheetName}
                          variant={selectedWorksheet === sheetName ? "default" : "outline"}
                          size="sm"
                          onClick={() => handleWorksheetChange(sheetName)}
                          className="justify-start text-left"
                          disabled={isProcessing || isParsingFile || isPolling}
                        >
                          {sheetName}
                        </Button>
                      ))}
                    </div>
                  </div>
                )}

                <Separator />

                {/* Action Buttons */}
                <div className="space-y-3">
                  <Button
                    onClick={handleSubmit}
                    disabled={isProcessing || isParsingFile || isPolling || !instanceUrl || !cookie || !xlFile}
                    className="w-full h-11 bg-gradient-to-r from-green-600 to-emerald-600 hover:from-green-700 hover:to-emerald-700 text-white font-medium"
                  >
                    {isProcessing || isPolling ? (
                      <>
                        <RefreshCw className="w-4 h-4 mr-2 animate-spin" />
                        {isPolling ? "Processing..." : "Starting..."}
                      </>
                    ) : (
                      <>
                        <Play className="w-4 h-4 mr-2" />
                        Start Processing
                      </>
                    )}
                  </Button>

                  <Button
                    onClick={resetForm}
                    variant="outline"
                    className="w-full h-11 border-slate-200 bg-transparent"
                    disabled={isProcessing && !canCancel}
                  >
                    <RefreshCw className="w-4 h-4 mr-2" />
                    Reset Form
                  </Button>
                </div>

                {/* Processing Info */}
                {(isProcessing || isPolling) && (
                  <div className="mt-4 p-4 bg-blue-50 rounded-lg border border-blue-200">
                    <div className="text-sm text-blue-800 space-y-2">
                      <div className="flex items-center">
                        <Clock className="w-4 h-4 mr-2" />
                        <span className="font-medium">Long-running process active</span>
                      </div>
                      <p className="text-xs text-blue-600">
                        This process can take up to 30 minutes. You can safely close this tab and return later.
                        {jobId && ` Job ID: ${jobId}`}
                      </p>
                    </div>
                  </div>
                )}
              </CardContent>
            </Card>
          </div>

          {/* Results Panel */}
          <div className="lg:col-span-2">
            <Tabs defaultValue="overview" className="space-y-6">
              <TabsList className="grid w-full grid-cols-3 bg-slate-100">
                <TabsTrigger value="overview" className="flex items-center">
                  <TrendingUp className="w-4 h-4 mr-2" />
                  Overview
                </TabsTrigger>
                <TabsTrigger value="preview" className="flex items-center">
                  <Eye className="w-4 h-4 mr-2" />
                  Preview
                </TabsTrigger>
                <TabsTrigger value="results" className="flex items-center">
                  <Activity className="w-4 h-4 mr-2" />
                  Results
                </TabsTrigger>
              </TabsList>

              <TabsContent value="preview">
                <Card className="shadow-lg border-0">
                  <CardHeader>
                    <CardTitle className="flex items-center justify-between">
                      <span className="flex items-center">
                        <FileText className="w-5 h-5 mr-2 text-green-600" />
                        Excel Data Preview
                      </span>
                      <div className="flex items-center space-x-2">
                        {hyperlinkCount > 0 && (
                          <Badge variant="outline" className="bg-blue-50 text-blue-700 border-blue-200">
                            <ExternalLink className="w-3 h-3 mr-1" />
                            {hyperlinkCount} Links
                          </Badge>
                        )}
                        {selectedWorksheet && (
                          <Badge variant="outline" className="bg-green-50 text-green-700 border-green-200">
                            {selectedWorksheet}
                          </Badge>
                        )}
                      </div>
                    </CardTitle>
                    <p className="text-sm text-slate-500">
                      Preview of your uploaded Excel file (first 5 rows) - Hyperlinks are preserved and clickable
                    </p>
                  </CardHeader>
                  <CardContent>
                    {isParsingFile ? (
                      <div className="text-center py-12">
                        <RefreshCw className="w-12 h-12 text-green-500 mx-auto mb-4 animate-spin" />
                        <p className="text-slate-500">Parsing Excel file and extracting hyperlinks...</p>
                        <p className="text-sm text-slate-400 mt-1">Please wait while we process your file</p>
                      </div>
                    ) : xlPreview.length > 0 ? (
                      <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                          <thead>
                            <tr className="bg-slate-50 rounded-lg">
                              {Object.keys(xlPreview[0] || {}).map((header, index) => (
                                <th
                                  key={index}
                                  className="px-4 py-3 text-left font-medium text-slate-700 border-b border-slate-200"
                                >
                                  {header}
                                </th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {xlPreview.map((row, index) => (
                              <tr key={index} className="border-b border-slate-100 hover:bg-slate-50">
                                {Object.values(row).map((cell, cellIndex) => (
                                  <td key={cellIndex} className="px-4 py-3 text-slate-600">
                                    {renderCellContent(cell)}
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      <div className="text-center py-12">
                        <FileText className="w-12 h-12 text-slate-300 mx-auto mb-4" />
                        <p className="text-slate-500">No file uploaded yet</p>
                        <p className="text-sm text-slate-400 mt-1">
                          Upload an Excel file to see the preview with hyperlinks
                        </p>
                      </div>
                    )}
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="overview" className="space-y-6">
                {/* Statistics Cards */}
                {apiResponse ? (
                  <>
                    <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                      <Card className="bg-gradient-to-br from-slate-50 to-slate-100 border-slate-200">
                        <CardContent className="pt-6">
                          <div className="flex items-center justify-between">
                            <div>
                              <p className="text-2xl font-bold text-slate-700">{stats.total}</p>
                              <p className="text-sm text-slate-500">Total Records</p>
                            </div>
                            <Users className="w-8 h-8 text-slate-400" />
                          </div>
                        </CardContent>
                      </Card>
                      <Card className="bg-gradient-to-br from-emerald-50 to-emerald-100 border-emerald-200">
                        <CardContent className="pt-6">
                          <div className="flex items-center justify-between">
                            <div>
                              <p className="text-2xl font-bold text-emerald-700">{stats.success}</p>
                              <p className="text-sm text-emerald-600">Successful</p>
                            </div>
                            <CheckCircle className="w-8 h-8 text-emerald-400" />
                          </div>
                        </CardContent>
                      </Card>
                      <Card className="bg-gradient-to-br from-amber-50 to-amber-100 border-amber-200">
                        <CardContent className="pt-6">
                          <div className="flex items-center justify-between">
                            <div>
                              <p className="text-2xl font-bold text-amber-700">{stats.partial}</p>
                              <p className="text-sm text-amber-600">Partial</p>
                            </div>
                            <AlertCircle className="w-8 h-8 text-amber-400" />
                          </div>
                        </CardContent>
                      </Card>
                      <Card className="bg-gradient-to-br from-red-50 to-red-100 border-red-200">
                        <CardContent className="pt-6">
                          <div className="flex items-center justify-between">
                            <div>
                              <p className="text-2xl font-bold text-red-700">{stats.failed}</p>
                              <p className="text-sm text-red-600">Failed</p>
                            </div>
                            <X className="w-8 h-8 text-red-400" />
                          </div>
                        </CardContent>
                      </Card>
                    </div>

                    {/* Progress Overview */}
                    <Card className="shadow-lg border-0">
                      <CardHeader>
                        <CardTitle className="flex items-center justify-between">
                          <span className="flex items-center">
                            <Activity className="w-5 h-5 mr-2 text-green-600" />
                            Processing Summary
                          </span>
                          <Badge className="bg-green-100 text-green-700">
                            {stats.total > 0 ? Math.round((stats.success / stats.total) * 100) : 0}% Success Rate
                          </Badge>
                        </CardTitle>
                      </CardHeader>
                      <CardContent className="space-y-4">
                        <div className="space-y-2">
                          <div className="flex justify-between text-sm">
                            <span className="text-slate-600">Overall Progress</span>
                            <span className="font-medium">
                              {stats.success}/{stats.total} completed
                            </span>
                          </div>
                          <Progress value={stats.total > 0 ? (stats.success / stats.total) * 100 : 0} className="h-2" />
                        </div>

                        <div className="grid grid-cols-3 gap-4 pt-4 border-t border-slate-100">
                          <div className="text-center">
                            <div className="text-lg font-semibold text-emerald-600">{stats.success}</div>
                            <div className="text-xs text-slate-500">Success</div>
                          </div>
                          <div className="text-center">
                            <div className="text-lg font-semibold text-amber-600">{stats.partial}</div>
                            <div className="text-xs text-slate-500">Partial</div>
                          </div>
                          <div className="text-center">
                            <div className="text-lg font-semibold text-red-600">{stats.failed}</div>
                            <div className="text-xs text-slate-500">Failed</div>
                          </div>
                        </div>

                        {/* File Processing Info */}
                        {apiResponse.file && (
                          <div className="pt-4 border-t border-slate-100">
                            <div className="text-sm text-slate-600 space-y-1">
                              <div className="flex justify-between">
                                <span>File:</span>
                                <span className="font-medium">{apiResponse.file.originalname}</span>
                              </div>
                              <div className="flex justify-between">
                                <span>Worksheet:</span>
                                <span className="font-medium">{apiResponse.file.worksheet}</span>
                              </div>
                              <div className="flex justify-between">
                                <span>Size:</span>
                                <span className="font-medium">{(apiResponse.file.size / 1024).toFixed(2)} KB</span>
                              </div>
                              <div className="flex justify-between">
                                <span>Hyperlinks:</span>
                                <span className="font-medium text-blue-600">{hyperlinkCount} found</span>
                              </div>
                            </div>
                          </div>
                        )}
                      </CardContent>
                    </Card>
                  </>
                ) : (
                  <Card className="shadow-lg border-0">
                    <CardContent className="py-16">
                      <div className="text-center">
                        <div className="w-16 h-16 bg-slate-100 rounded-full flex items-center justify-center mx-auto mb-4">
                          <Database className="w-8 h-8 text-slate-400" />
                        </div>
                        <h3 className="text-lg font-medium text-slate-900 mb-2">Ready to Process</h3>
                        <p className="text-slate-500 mb-6">
                          Upload an Excel file and configure your settings to begin processing
                        </p>
                        <div className="flex items-center justify-center space-x-6 text-sm text-slate-400">
                          <div className="flex items-center">
                            <Clock className="w-4 h-4 mr-2" />
                            Long-Running Support
                          </div>
                          <div className="flex items-center">
                            <Shield className="w-4 h-4 mr-2" />
                            Secure Upload
                          </div>
                          <div className="flex items-center">
                            <ExternalLink className="w-4 h-4 mr-2" />
                            Hyperlink Support
                          </div>
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                )}
              </TabsContent>

              <TabsContent value="results">
                <Card className="shadow-lg border-0">
                  <CardHeader>
                    <CardTitle className="flex items-center justify-between">
                      <span className="flex items-center">
                        <Activity className="w-5 h-5 mr-2 text-green-600" />
                        Processing Results
                      </span>
                      {apiResponse && (
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => {
                            const dataStr = JSON.stringify(apiResponse, null, 2)
                            const dataBlob = new Blob([dataStr], { type: "application/json" })
                            const url = URL.createObjectURL(dataBlob)
                            const link = document.createElement("a")
                            link.href = url
                            link.download = `processing-results-${new Date().toISOString().split("T")[0]}.json`
                            link.click()
                            URL.revokeObjectURL(url)
                          }}
                        >
                          <Download className="w-4 h-4 mr-2" />
                          Export
                        </Button>
                      )}
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    {apiResponse ? (
                      <div className="space-y-4">
                        {/* File Info */}
                        <Card className="bg-green-50 border-green-200">
                          <CardContent className="pt-4">
                            <div className="flex items-center justify-between">
                              <div className="flex items-center space-x-3">
                                <FileText className="w-8 h-8 text-green-600" />
                                <div>
                                  <p className="font-medium text-green-900">
                                    {apiResponse.file?.originalname || apiResponse.file?.filename}
                                  </p>
                                  <p className="text-sm text-green-700">
                                    {apiResponse.file?.size ? `${(apiResponse.file.size / 1024).toFixed(2)} KB` : ""} •
                                    Processed successfully
                                  </p>
                                </div>
                              </div>
                              <Badge className="bg-green-100 text-green-800">
                                {apiResponse.success ? "Completed" : "Failed"}
                              </Badge>
                            </div>
                          </CardContent>
                        </Card>

                        {/* Detailed Results */}
                        <div className="space-y-3 max-h-96 overflow-y-auto">
                          {apiResponse.processedData?.map((record, index) => (
                            <Card
                              key={index}
                              className="border-l-4 border-l-green-500 hover:shadow-md transition-shadow"
                            >
                              <CardContent className="pt-4">
                                <div className="flex items-start justify-between mb-3">
                                  <div className="flex-1">
                                    <p className="font-medium text-slate-800 text-sm">
                                      Company: {record.Company_GSID?.substring(0, 30)}...
                                    </p>
                                    <p className="text-xs text-slate-500 mt-1">Email: {record.Invite_Email}</p>
                                  </div>
                                  <Badge className={`${getStatusColor(record.status)} flex items-center gap-1`}>
                                    {getStatusIcon(record.status)}
                                    {record.status}
                                  </Badge>
                                </div>

                                <div className="space-y-1">
                                  {record.messages?.map((message, msgIndex) => (
                                    <div key={msgIndex} className="flex items-center text-xs">
                                      <div className="w-1.5 h-1.5 bg-green-400 rounded-full mr-2 flex-shrink-0"></div>
                                      <span className="text-slate-600">{message}</span>
                                    </div>
                                  ))}
                                </div>

                                {record.Video_URL && (
                                  <div className="mt-3 pt-3 border-t border-slate-100">
                                    <a
                                      href={record.Video_URL}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className="text-xs text-green-600 hover:text-green-800 flex items-center font-medium"
                                    >
                                      <Play className="w-3 h-3 mr-1" />
                                      View Generated Video
                                    </a>
                                  </div>
                                )}
                              </CardContent>
                            </Card>
                          ))}
                        </div>
                      </div>
                    ) : (
                      <div className="text-center py-12">
                        <Activity className="w-12 h-12 text-slate-300 mx-auto mb-4" />
                        <p className="text-slate-500">No results yet</p>
                        <p className="text-sm text-slate-400 mt-1">Process your data to see detailed results</p>
                      </div>
                    )}
                  </CardContent>
                </Card>
              </TabsContent>
            </Tabs>
          </div>
        </div>
      </div>

      {/* Cookie Modal */}
      {showCookieModal && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <Card className="w-full max-w-2xl max-h-[80vh] overflow-hidden shadow-2xl">
            <CardHeader className="border-b border-slate-200">
              <CardTitle className="flex items-center justify-between">
                <span className="flex items-center">
                  <Shield className="w-5 h-5 mr-2 text-green-600" />
                  Authentication Cookie
                </span>
                <Button variant="ghost" size="sm" onClick={() => setShowCookieModal(false)}>
                  <X className="w-4 h-4" />
                </Button>
              </CardTitle>
            </CardHeader>
            <CardContent className="p-0">
              <div className="bg-slate-50 p-6 max-h-96 overflow-y-auto">
                <pre className="text-xs text-slate-700 whitespace-pre-wrap break-all font-mono">
                  {cookie || "No cookie content available"}
                </pre>
              </div>
            </CardContent>
          </Card>
        </div>
      )}
    </div>
  )
}

export default Unit4XlDashboard
