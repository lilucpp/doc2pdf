cmake_minimum_required(VERSION 3.10)
project(wps2pdf LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

#if(CMAKE_SIZEOF_VOID_P EQUAL 8)
#    message(FATAL_ERROR "This project must be built as 32-bit (x86). Use -A Win32 or set CMAKE_GENERATOR_PLATFORM=Win32")
#endif()

if(MSVC)
    set(CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE} /MT /utf-8")
    set(CMAKE_CXX_FLAGS_DEBUG "${CMAKE_CXX_FLAGS_DEBUG} /MTd /utf-8")
endif()

add_executable(wps2pdf wps2pdf.cpp)
target_link_libraries(wps2pdf ole32 oleaut32)
