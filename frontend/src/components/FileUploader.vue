<template>
  <div class="w-full max-w-4xl mx-auto px-4 py-8 relative overflow-hidden">
    <!-- Floating Animated Illustrations -->
    <div class="absolute inset-0 pointer-events-none">
      <!-- Floating File Icons -->
      <div class="absolute top-20 left-10 animate-float opacity-20">
        <svg class="w-8 h-8 text-blue-400" fill="currentColor" viewBox="0 0 20 20">
          <path fill-rule="evenodd" d="M4 4a2 2 0 012-2h4.586A2 2 0 0112 2.586L15.414 6A2 2 0 0116 7.414V16a2 2 0 01-2 2H6a2 2 0 01-2-2V4z" clip-rule="evenodd" />
        </svg>
      </div>
      
      <div class="absolute top-40 right-16 animate-float-delayed opacity-20">
        <svg class="w-6 h-6 text-purple-400" fill="currentColor" viewBox="0 0 20 20">
          <path d="M4 3a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V5a2 2 0 00-2-2H4zm12 12H4l4-8 3 6 2-4 3 6z" />
        </svg>
      </div>
      
      <div class="absolute top-60 left-20 animate-bounce-slow opacity-20">
        <svg class="w-7 h-7 text-green-400" fill="currentColor" viewBox="0 0 20 20">
          <path fill-rule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm3.293-7.707a1 1 0 011.414 0L9 10.586V3a1 1 0 112 0v7.586l1.293-1.293a1 1 0 111.414 1.414L10 14.414l-3.707-3.707a1 1 0 010-1.414z" clip-rule="evenodd" />
        </svg>
      </div>
      
      <div class="absolute bottom-40 right-8 animate-pulse opacity-20">
        <svg class="w-5 h-5 text-indigo-400" fill="currentColor" viewBox="0 0 20 20">
          <path fill-rule="evenodd" d="M4 2a1 1 0 011 1v2.101a7.002 7.002 0 0111.601 2.566 1 1 0 11-1.885.666A5.002 5.002 0 005.999 7H9a1 1 0 010 2H4a1 1 0 01-1-1V3a1 1 0 011-1z" clip-rule="evenodd" />
        </svg>
      </div>
    </div>

    <!-- Header Section -->
    <div class="text-center mb-8 relative z-10">
      <div class="mb-6">
        <!-- Animated Logo/Icon -->
        <div class="inline-flex items-center justify-center w-12 h-12 sm:w-16 sm:h-16 bg-gradient-to-r from-blue-600 to-purple-600 rounded-full mb-4 animate-spin-slow shadow-lg">
          <svg class="w-6 h-6 sm:w-8 sm:h-8 text-white" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M4 2a1 1 0 011 1v2.101a7.002 7.002 0 0111.601 2.566 1 1 0 11-1.885.666A5.002 5.002 0 005.999 7H9a1 1 0 010 2H4a1 1 0 01-1-1V3a1 1 0 011-1zm.008 9.057a1 1 0 011.276.61A5.002 5.002 0 0014.001 13H11a1 1 0 110-2h5a1 1 0 011 1v5a1 1 0 11-2 0v-2.101a7.002 7.002 0 01-11.601-2.566 1 1 0 01.61-1.276z" clip-rule="evenodd" />
          </svg>
        </div>
      </div>
      
      <!-- Animated Title -->
      <h1 class="text-2xl sm:text-3xl lg:text-4xl font-bold text-gray-900 mb-3 px-4">
        <span class="bg-gradient-to-r from-blue-600 via-purple-600 to-blue-800 bg-clip-text text-transparent animate-gradient-x bg-300% inline-block">
          Universal File Converter Pro
        </span>
      </h1>
      
      <!-- Subtitle with typing animation -->
      <p class="text-lg sm:text-xl text-gray-600 mb-2 px-4 animate-fade-in-up">
        Convert any file to any format instantly
      </p>
      
      <!-- Description -->
      <p class="text-sm sm:text-base text-gray-500 max-w-2xl mx-auto px-4 leading-relaxed animate-fade-in-up-delayed">
        Professional file conversion supporting 22+ formats. Convert, merge, or batch process your documents, images, spreadsheets, and presentations with ease.
      </p>
      
      <!-- Privacy & Features Tags -->
      <div class="flex flex-wrap justify-center gap-2 sm:gap-3 mt-6 mb-2 px-4 max-w-3xl mx-auto">
        <div class="group relative">
          <div class="flex items-center gap-1.5 px-3 py-1.5 bg-green-50 hover:bg-green-100 text-green-700 rounded-full text-xs font-medium border border-green-200 transition-all duration-200 hover:scale-105 cursor-default">
            <svg class="w-3 h-3" fill="currentColor" viewBox="0 0 20 20">
              <path fill-rule="evenodd" d="M2.166 4.999A11.954 11.954 0 0010 1.944 11.954 11.954 0 0017.834 5c.11.65.166 1.32.166 2.001 0 5.225-3.34 9.67-8 11.317C5.34 16.67 2 12.225 2 7c0-.682.057-1.35.166-2.001zm11.541 3.708a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd" />
            </svg>
            <span>100% Private</span>
          </div>
        </div>
        
        <div class="group relative">
          <div class="flex items-center gap-1.5 px-3 py-1.5 bg-blue-50 hover:bg-blue-100 text-blue-700 rounded-full text-xs font-medium border border-blue-200 transition-all duration-200 hover:scale-105 cursor-default">
            <svg class="w-3 h-3" fill="currentColor" viewBox="0 0 20 20">
              <path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clip-rule="evenodd" />
            </svg>
            <span>No Data Stored</span>
          </div>
        </div>
        
        <div class="group relative">
          <div class="flex items-center gap-1.5 px-3 py-1.5 bg-purple-50 hover:bg-purple-100 text-purple-700 rounded-full text-xs font-medium border border-purple-200 transition-all duration-200 hover:scale-105 cursor-default">
            <svg class="w-3 h-3" fill="currentColor" viewBox="0 0 20 20">
              <path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
            <span>Free Forever</span>
          </div>
        </div>
        
        <div class="group relative">
          <div class="flex items-center gap-1.5 px-3 py-1.5 bg-gray-50 hover:bg-gray-100 text-gray-700 rounded-full text-xs font-medium border border-gray-200 transition-all duration-200 hover:scale-105 cursor-default">
            <svg class="w-3 h-3" fill="currentColor" viewBox="0 0 20 20">
              <path fill-rule="evenodd" d="M3 4a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm0 4a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm0 4a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm0 4a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1z" clip-rule="evenodd" />
            </svg>
            <span>No Ads</span>
          </div>
        </div>
        
        <div class="group relative">
          <div class="flex items-center gap-1.5 px-3 py-1.5 bg-orange-50 hover:bg-orange-100 text-orange-700 rounded-full text-xs font-medium border border-orange-200 transition-all duration-200 hover:scale-105 cursor-default">
            <svg class="w-3 h-3" fill="currentColor" viewBox="0 0 20 20">
              <path fill-rule="evenodd" d="M18 8a6 6 0 01-7.743 5.743L10 14l-1 1-1 1H6v2H2v-4l4.257-4.257A6 6 0 1118 8zm-6-4a1 1 0 100 2 2 2 0 012 2 1 1 0 102 0 4 4 0 00-4-4z" clip-rule="evenodd" />
            </svg>
            <span>No Sign-up</span>
          </div>
        </div>
      </div>
    </div>

    <div
      class="w-full max-w-2xl mx-auto p-4 sm:p-6 lg:p-8 border-2 border-dashed border-gray-300 rounded-xl text-center bg-gradient-to-br from-blue-50 to-purple-50 hover:from-blue-100 hover:to-purple-100 hover:border-blue-400 transition-all duration-300 shadow-sm relative z-10"
      @dragover.prevent
      @dragenter.prevent="dragActive = true"
      @dragleave.prevent="dragActive = false"
      @drop.prevent="handleDrop"
      :class="{ 'border-blue-400 bg-gradient-to-br from-blue-100 to-purple-100': dragActive }"
    >
    <!-- Upload Icon -->
    <div class="mb-4">
      <svg class="mx-auto h-10 w-10 sm:h-12 sm:w-12 text-gray-400" stroke="currentColor" fill="none" viewBox="0 0 48 48">
        <path d="M28 8H12a4 4 0 00-4 4v20m32-12v8m0 0v8a4 4 0 01-4 4H12a4 4 0 01-4-4v-4m32-4l-3.172-3.172a4 4 0 00-5.656 0L28 28M8 32l9.172-9.172a4 4 0 015.656 0L28 28m0 0l4 4m4-24h8m-4-4v8m-12 4h.02" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" />
      </svg>
    </div>
    
    <h3 class="text-lg sm:text-xl font-semibold text-gray-800 mb-2">Upload Your Files</h3>
    <p class="text-sm sm:text-base text-gray-600 mb-6 px-2">
      Drag & drop files here or choose an option below
    </p>

    <!-- Hidden file inputs -->
    <input type="file" ref="singleFileInput" @change="handleFiles" class="hidden" accept="image/*,.pdf,.docx,.doc,.txt,.rtf,.odt,.xlsx,.xls,.csv,.ods,.pptx,.ppt,.odp,.svg,.heic" />
    <input type="file" ref="multipleFileInput" @change="handleFiles" multiple class="hidden" accept="image/*,.pdf,.docx,.doc,.txt,.rtf,.odt,.xlsx,.xls,.csv,.ods,.pptx,.ppt,.odp,.svg,.heic" />
    
    <!-- File Selection Buttons -->
    <div class="flex flex-col gap-3 sm:flex-row sm:gap-3 justify-center mb-6">
      <button
        @click="$refs.singleFileInput.click()"
        class="w-full sm:w-auto inline-flex items-center justify-center px-4 sm:px-6 py-3 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 transition-colors shadow-md text-sm sm:text-base"
      >
        <svg class="w-4 h-4 sm:w-5 sm:h-5 mr-2" fill="currentColor" viewBox="0 0 20 20">
          <path fill-rule="evenodd" d="M8 4a3 3 0 00-3 3v4a5 5 0 0010 0V7a1 1 0 112 0v4a7 7 0 11-14 0V7a5 5 0 0110 0v4a3 3 0 11-6 0V7a1 1 0 012 0v4a1 1 0 102 0V7a3 3 0 00-3-3z" clip-rule="evenodd" />
        </svg>
        Select Single File
      </button>
      
      <button
        @click="$refs.multipleFileInput.click()"
        class="w-full sm:w-auto inline-flex items-center justify-center px-4 sm:px-6 py-3 bg-green-600 text-white font-medium rounded-lg hover:bg-green-700 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-offset-2 transition-colors shadow-md text-sm sm:text-base"
      >
        <svg class="w-4 h-4 sm:w-5 sm:h-5 mr-2" fill="currentColor" viewBox="0 0 20 20">
          <path d="M5.5 13a3.5 3.5 0 01-.369-6.98 4 4 0 117.753-1.977A4.5 4.5 0 1113.5 13H11V9.413l1.293 1.293a1 1 0 001.414-1.414l-3-3a1 1 0 00-1.414 0l-3 3a1 1 0 001.414 1.414L9 9.413V13H5.5z" />
        </svg>
        Select Multiple Files
      </button>
    </div>


    <div v-if="files.length" class="mb-6">
      <h3 class="text-lg font-semibold text-gray-800 mb-4 text-center">Selected Files</h3>
      <div class="space-y-3">
        <div 
          v-for="(file, index) in files" 
          :key="index" 
          class="p-4 bg-white rounded-lg shadow-sm border border-gray-200 hover:shadow-md transition-shadow"
        >
          <div class="flex items-center gap-4">
            <!-- File Icon/Preview -->
            <div class="flex-shrink-0">
              <img 
                v-if="isImage(file)" 
                :src="file.preview" 
                class="w-12 h-12 sm:w-16 sm:h-16 object-cover rounded-lg border border-gray-200" 
                :alt="file.name"
              />
              <div 
                v-else 
                class="w-12 h-12 sm:w-16 sm:h-16 bg-gray-100 rounded-lg flex items-center justify-center border border-gray-200"
              >
                <svg class="w-6 h-6 sm:w-8 sm:h-8 text-gray-400" fill="currentColor" viewBox="0 0 20 20">
                  <path fill-rule="evenodd" d="M4 4a2 2 0 012-2h4.586A2 2 0 0112 2.586L15.414 6A2 2 0 0116 7.414V16a2 2 0 01-2 2H6a2 2 0 01-2-2V4zm2 6a1 1 0 011-1h6a1 1 0 110 2H7a1 1 0 01-1-1zm1 3a1 1 0 100 2h6a1 1 0 100-2H7z" clip-rule="evenodd" />
                </svg>
              </div>
            </div>
            
            <!-- File Info -->
            <div class="flex-1 min-w-0">
              <p class="text-sm sm:text-base font-medium text-gray-900 truncate">{{ file.name }}</p>
              <p class="text-xs sm:text-sm text-gray-500 mt-1">{{ formatFileType(file.type) }} • {{ formatFileSize(file.file.size) }}</p>
            </div>
            
            <!-- Remove Button -->
            <button 
              @click="removeFile(index)"
              class="flex-shrink-0 text-red-400 hover:text-red-600 transition-colors p-2 hover:bg-red-50 rounded-lg"
              title="Remove file"
            >
              <svg class="w-5 h-5" fill="currentColor" viewBox="0 0 20 20">
                <path fill-rule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clip-rule="evenodd" />
              </svg>
            </button>
          </div>
        </div>
      </div>
    </div>

    <div v-if="files.length" class="mb-6">
      <label class="block text-base sm:text-lg font-semibold text-gray-800 mb-3 text-center">
        Convert {{ files.length }} file{{ files.length > 1 ? 's' : '' }} to:
      </label>
      
      <div class="relative">
        <select 
          v-model="selectedFormat" 
          class="w-full px-3 sm:px-4 py-2 sm:py-3 border-2 border-gray-300 rounded-lg sm:rounded-xl shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white text-gray-900 font-medium text-sm sm:text-lg appearance-none cursor-pointer hover:border-blue-400 transition-colors"
          :disabled="isConverting"
        >
          <optgroup label="Document Formats" class="font-semibold">
            <option value="pdf" class="py-2">PDF - Portable Document Format</option>
            <option value="docx" class="py-2">DOCX - Microsoft Word Document</option>
            <option value="doc" class="py-2">DOC - Legacy Word Document</option>
            <option value="txt" class="py-2">TXT - Plain Text File</option>
            <option value="rtf" class="py-2">RTF - Rich Text Format</option>
            <option value="odt" class="py-2">ODT - OpenDocument Text</option>
          </optgroup>
          
          <optgroup label="Image Formats" class="font-semibold">
            <option value="png" class="py-2">PNG - High Quality with Transparency</option>
            <option value="jpg" class="py-2">JPG - Standard Photo Format</option>
            <option value="jpeg" class="py-2">JPEG - Compressed Image Format</option>
            <option value="webp" class="py-2">WebP - Modern Web Format</option>
            <option value="bmp" class="py-2">BMP - Bitmap Image</option>
            <option value="tiff" class="py-2">TIFF - High Quality Archive Format</option>
            <option value="gif" class="py-2">GIF - Animated/Simple Graphics</option>
            <option value="svg" class="py-2">SVG - Scalable Vector Graphics</option>
            <option value="ico" class="py-2">ICO - Icon Format</option>
            <option value="heic" class="py-2">HEIC - Apple High Efficiency</option>
          </optgroup>
          
          <optgroup label="Spreadsheet Formats" class="font-semibold">
            <option value="xlsx" class="py-2">XLSX - Excel Spreadsheet</option>
            <option value="xls" class="py-2">XLS - Legacy Excel Format</option>
            <option value="csv" class="py-2">CSV - Comma Separated Values</option>
            <option value="ods" class="py-2">ODS - OpenDocument Spreadsheet</option>
          </optgroup>
          
          <optgroup label="Presentation Formats" class="font-semibold">
            <option value="pptx" class="py-2">PPTX - PowerPoint Presentation</option>
            <option value="ppt" class="py-2">PPT - Legacy PowerPoint</option>
            <option value="odp" class="py-2">ODP - OpenDocument Presentation</option>
          </optgroup>
          
          <optgroup label="Archive Options" v-if="files.length > 1" class="font-semibold">
            <option value="zip" class="py-2">ZIP - Individual Converted Files</option>
          </optgroup>
        </select>
        
        <!-- Custom dropdown arrow -->
        <div class="absolute inset-y-0 right-0 flex items-center px-2 sm:px-3 pointer-events-none">
          <svg class="w-4 h-4 sm:w-5 sm:h-5 text-gray-400" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z" clip-rule="evenodd" />
          </svg>
        </div>
      </div>
      
      <!-- Enhanced conversion info with better styling -->
      <div class="mt-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
        <div class="text-xs sm:text-sm text-blue-800 font-medium mb-1 text-center">
          <span v-if="files.length === 1">
            Converting: {{ files[0].name }} → {{ selectedFormat.toUpperCase() }}
          </span>
          <span v-else-if="selectedFormat === 'zip'">
            Each file will be converted individually and packaged in a ZIP archive
          </span>
          <span v-else>
            {{ files.length }} files will be merged into one {{ selectedFormat.toUpperCase() }} document
          </span>
        </div>
        
        <!-- File count and total size with icons -->
        <div class="text-xs sm:text-sm text-blue-600 flex items-center justify-center gap-2 sm:gap-4 flex-wrap">
          <span class="flex items-center gap-1">
            {{ files.length }} file{{ files.length > 1 ? 's' : '' }}
          </span>
          <span class="flex items-center gap-1">
            💾 {{ getTotalFileSize() }}
          </span>
        </div>
      </div>
    </div>

    <!-- Conversion Progress -->
    <div v-if="isConverting" class="mb-6">
      <div class="bg-blue-50 border border-blue-200 rounded-lg p-4">
        <div class="flex items-center mb-2">
          <svg class="animate-spin h-5 w-5 text-blue-600 mr-2" fill="none" viewBox="0 0 24 24">
            <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
            <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
          </svg>
          <span class="text-blue-800 font-medium">Converting files...</span>
        </div>
        
        <!-- Progress Bar -->
        <div class="w-full bg-blue-200 rounded-full h-2 mb-2">
          <div 
            class="bg-blue-600 h-2 rounded-full transition-all duration-300 ease-out" 
            :style="{ width: conversionProgress + '%' }"
          ></div>
        </div>
        
        <p class="text-sm text-blue-700">
          Processing {{ files.length }} files into {{ selectedFormat.toUpperCase() }} format...
        </p>
      </div>
    </div>

    <!-- Action Buttons -->
    <div v-if="files.length" class="flex flex-col gap-4">
      <!-- Convert Button (shown when not converting and no download) -->
      <button
        v-if="!isConverting && !downloadUrl"
        @click="convertFiles"
        class="group w-full px-4 sm:px-6 py-3 sm:py-4 bg-gray-900 hover:bg-black text-white font-medium text-sm sm:text-base rounded-lg transition-all duration-200 ease-out hover:shadow-md focus:outline-none focus:ring-2 focus:ring-gray-300 focus:ring-offset-2 active:scale-[0.98]"
      >
        <div class="flex items-center justify-center gap-2 sm:gap-3">
          <svg class="w-4 h-4 sm:w-5 sm:h-5 transition-transform duration-200 group-hover:rotate-180" fill="currentColor" viewBox="0 0 20 20">
            <path fill-rule="evenodd" d="M4 2a1 1 0 011 1v2.101a7.002 7.002 0 0111.601 2.566 1 1 0 11-1.885.666A5.002 5.002 0 005.999 7H9a1 1 0 010 2H4a1 1 0 01-1-1V3a1 1 0 011-1zm.008 9.057a1 1 0 011.276.61A5.002 5.002 0 0014.001 13H11a1 1 0 110-2h5a1 1 0 011 1v5a1 1 0 11-2 0v-2.101a7.002 7.002 0 01-11.601-2.566 1 1 0 01.61-1.276z" clip-rule="evenodd" />
          </svg>
          <span class="tracking-wide">
            Convert {{ files.length }} file<span v-if="files.length > 1">s</span> to {{ selectedFormat.toUpperCase() }}
          </span>
        </div>
      </button>
      
      <!-- Download Button (shown when conversion is complete) -->
      <div v-if="!isConverting && downloadUrl" class="space-y-3">
        <!-- Success Message -->
        <div class="bg-green-50 border border-green-200 rounded-lg p-4">
          <div class="flex items-center">
            <svg class="w-5 h-5 text-green-600 mr-2" fill="currentColor" viewBox="0 0 20 20">
              <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd" />
            </svg>
            <span class="text-green-800 font-medium">Conversion completed successfully!</span>
          </div>
        </div>
        
        <!-- Download Button -->
        <a
          :href="downloadUrl"
          :download="convertedFileName"
          class="group w-full px-4 sm:px-6 py-3 sm:py-4 bg-slate-700 hover:bg-slate-800 text-slate-50 font-medium text-sm sm:text-base rounded-lg transition-all duration-200 ease-out hover:shadow-md focus:outline-none focus:ring-2 focus:ring-slate-300 focus:ring-offset-2 active:scale-[0.98] flex items-center justify-center animate-pulse hover:animate-none"
        >
          <div class="flex items-center gap-2 sm:gap-3">
            <svg class="w-4 h-4 sm:w-5 sm:h-5 transition-transform duration-200 group-hover:translate-y-0.5" fill="currentColor" viewBox="0 0 20 20">
              <path fill-rule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zm3.293-7.707a1 1 0 011.414 0L9 10.586V3a1 1 0 112 0v7.586l1.293-1.293a1 1 0 111.414 1.414L10 14.414l-3.707-3.707a1 1 0 010-1.414z" clip-rule="evenodd" />
            </svg>
            <span class="tracking-wide">Download {{ convertedFileName }}</span>
            <div class="w-1 h-1 bg-slate-100 rounded-full animate-ping ml-1"></div>
          </div>
        </a>
        
        <!-- Reset Button -->
        <button
          @click="resetConverter"
          class="group w-full px-4 sm:px-6 py-2.5 sm:py-3 bg-gray-100 hover:bg-gray-200 text-gray-700 hover:text-gray-900 font-medium text-sm sm:text-base rounded-lg transition-all duration-200 ease-out focus:outline-none focus:ring-2 focus:ring-gray-300 focus:ring-offset-2 active:scale-[0.98] flex items-center justify-center"
        >
          <div class="flex items-center gap-2 sm:gap-3">
            <svg class="w-4 h-4 sm:w-5 sm:h-5 transition-transform duration-200 group-hover:rotate-180" fill="currentColor" viewBox="0 0 20 20">
              <path fill-rule="evenodd" d="M4 2a1 1 0 011 1v2.101a7.002 7.002 0 0111.601 2.566 1 1 0 11-1.885.666A5.002 5.002 0 005.999 7H9a1 1 0 010 2H4a1 1 0 01-1-1V3a1 1 0 011-1zm.008 9.057a1 1 0 011.276.61A5.002 5.002 0 0014.001 13H11a1 1 0 110-2h5a1 1 0 011 1v5a1 1 0 11-2 0v-2.101a7.002 7.002 0 01-11.601-2.566 1 1 0 01.61-1.276z" clip-rule="evenodd" />
            </svg>
            <span class="tracking-wide">Convert New Files</span>
          </div>
        </button>
      </div>
    </div>
    </div>
    
    <!-- Footer Section -->
    <footer class="mt-20 py-6 border-t border-gray-100">
      <div class="text-center">
        <p class="text-gray-400 text-xs">
          © 2025 Built by 
          <a 
            href="https://www.shahinasareen.tech" 
            target="_blank" 
            rel="noopener noreferrer"
            class="text-gray-600 hover:text-gray-800 font-medium transition-colors"
          >
            Shahina Sareen
          </a>
          • All rights reserved
        </p>
      </div>
    </footer>
  </div>
</template><script>
export default {
  data() {
    return {
      files: [],
      selectedFormat: "pdf",
      downloadUrl: "",
      isConverting: false,
      dragActive: false,
      conversionProgress: 0,
      convertedFileName: "",
    };
  },
  methods: {
    handleFiles(event) {
      const selected = Array.from(event.target.files);
      this.setupPreview(selected);
    },
    handleDrop(event) {
      this.dragActive = false;
      const dropped = Array.from(event.dataTransfer.files);
      this.setupPreview(dropped);
    },
    setupPreview(fileList) {
      this.files = fileList.map((file) => ({
        file,
        name: file.name,
        type: file.type,
        preview: this.isImage(file) ? URL.createObjectURL(file) : null,
      }));
      this.downloadUrl = "";
    },
    removeFile(index) {
      // Clean up preview URL if it exists
      if (this.files[index].preview) {
        URL.revokeObjectURL(this.files[index].preview);
      }
      this.files.splice(index, 1);
      this.downloadUrl = "";
    },
    isImage(file) {
      return ["image/png", "image/jpeg", "image/jpg", "image/svg+xml"].includes(file.type);
    },
    formatFileType(mimeType) {
      const typeMap = {
        'image/png': 'PNG Image',
        'image/jpeg': 'JPEG Image', 
        'image/jpg': 'JPG Image',
        'image/svg+xml': 'SVG Vector',
        'application/pdf': 'PDF Document',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word Document'
      };
      return typeMap[mimeType] || mimeType.split('/')[1]?.toUpperCase() || 'Unknown';
    },
    formatFileSize(bytes) {
      if (bytes === 0) return '0 Bytes';
      const k = 1024;
      const sizes = ['Bytes', 'KB', 'MB', 'GB'];
      const i = Math.floor(Math.log(bytes) / Math.log(k));
      return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    },
    getTotalFileSize() {
      const totalBytes = this.files.reduce((sum, file) => sum + file.file.size, 0);
      return this.formatFileSize(totalBytes);
    },
    resetConverter() {
      // Clean up preview URLs
      this.files.forEach(file => {
        if (file.preview) {
          URL.revokeObjectURL(file.preview);
        }
      });
      
      // Clean up download URL
      if (this.downloadUrl) {
        URL.revokeObjectURL(this.downloadUrl);
      }
      
      // Reset state
      this.files = [];
      this.downloadUrl = "";
      this.conversionProgress = 0;
      this.convertedFileName = "";
      this.isConverting = false;
    },
    async convertFiles() {
      if (!this.files.length) return alert("Please select files!");
      
      this.isConverting = true;
      this.conversionProgress = 0;
      
      // Simulate progress updates
      const progressInterval = setInterval(() => {
        if (this.conversionProgress < 90) {
          this.conversionProgress += Math.random() * 20;
        }
      }, 200);

      const formData = new FormData();
      this.files.forEach((f) => formData.append("file", f.file));

      try {
        const res = await fetch(`http://127.0.0.1:5000/convert?format=${this.selectedFormat}`, {
          method: "POST",
          body: formData,
        });
        
        if (!res.ok) {
          const errorData = await res.json();
          throw new Error(errorData.error || "Conversion failed");
        }

        // Complete progress
        this.conversionProgress = 100;
        
        const blob = await res.blob();
        this.downloadUrl = URL.createObjectURL(blob);
        
        // Generate filename based on format and file count
        const timestamp = new Date().toISOString().slice(0, 10);
        
        if (this.selectedFormat === 'zip') {
          this.convertedFileName = `converted_files_${timestamp}.zip`;
        } else if (this.files.length === 1) {
          // Single file conversion
          const baseName = this.files[0].name.split('.')[0];
          this.convertedFileName = `${baseName}.${this.selectedFormat}`;
        } else if (['pdf', 'docx'].includes(this.selectedFormat)) {
          // Multiple files merged into one document
          this.convertedFileName = `merged_document_${timestamp}.${this.selectedFormat}`;
        } else {
          // Multiple files converted to same format (will be zipped)
          this.convertedFileName = `converted_to_${this.selectedFormat}_${timestamp}.zip`;
        }
        
      } catch (err) {
        alert(`Conversion failed: ${err.message}`);
        this.conversionProgress = 0;
      } finally {
        clearInterval(progressInterval);
        this.isConverting = false;
      }
    },
  },
};
</script>

<style scoped>
/* Custom Animations */
@keyframes float {
  0%, 100% { transform: translateY(0px); }
  50% { transform: translateY(-20px); }
}

@keyframes float-delayed {
  0%, 100% { transform: translateY(0px); }
  50% { transform: translateY(-15px); }
}

@keyframes bounce-slow {
  0%, 100% { transform: translateY(0); }
  50% { transform: translateY(-10px); }
}

@keyframes spin-slow {
  from { transform: rotate(0deg); }
  to { transform: rotate(360deg); }
}

@keyframes gradient-x {
  0%, 100% { background-position: 0% 50%; }
  50% { background-position: 100% 50%; }
}

@keyframes fade-in-up {
  from { 
    opacity: 0; 
    transform: translateY(20px); 
  }
  to { 
    opacity: 1; 
    transform: translateY(0); 
  }
}

@keyframes fade-in-up-delayed {
  from { 
    opacity: 0; 
    transform: translateY(20px); 
  }
  to { 
    opacity: 1; 
    transform: translateY(0); 
  }
}

/* Animation Classes */
.animate-float {
  animation: float 3s ease-in-out infinite;
}

.animate-float-delayed {
  animation: float-delayed 4s ease-in-out infinite;
}

.animate-bounce-slow {
  animation: bounce-slow 2s ease-in-out infinite;
}

.animate-spin-slow {
  animation: spin-slow 8s linear infinite;
}

.animate-gradient-x {
  animation: gradient-x 4s ease infinite;
}

.bg-300\% {
  background-size: 300% 300%;
}

.animate-fade-in-up {
  animation: fade-in-up 0.8s ease-out;
}

.animate-fade-in-up-delayed {
  animation: fade-in-up-delayed 1.2s ease-out;
}

/* Hover effects */
.group:hover .group-hover\:animate-bounce {
  animation: bounce 1s infinite;
}
</style>