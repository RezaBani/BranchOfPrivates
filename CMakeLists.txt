cmake_minimum_required(VERSION 3.5)

project(BranchOfPrivates VERSION 1.0 LANGUAGES CXX)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(APP_ICON_RESOURCE_WINDOWS "${CMAKE_CURRENT_SOURCE_DIR}/icon.rc")

find_package(QT NAMES Qt6 Qt5 REQUIRED COMPONENTS Widgets)
find_package(Qt${QT_VERSION_MAJOR} REQUIRED COMPONENTS Widgets)

set(PROJECT_SOURCES
    main.cpp
    mainwindow.cpp
    mainwindow.h
    mainwindow.ui
    addeditdialog.h
    addeditdialog.cpp
    addeditdialog.ui
    soldier.h
    soldier.cpp
    optiondialog.h
    optiondialog.cpp
    optiondialog.ui
    letterdialog.h
    letterdialog.cpp
    letterdialog.ui
    filterdialog.h
    filterdialog.cpp
    filterdialog.ui
    soldierfiles.h
    soldierfiles.cpp
    datecalc.h
    datecalc.cpp
    orderdialog.h
    orderdialog.cpp
    orderdialog.ui
    statisticsdialog.h
    statisticsdialog.cpp
    statisticsdialog.ui
    resources.qrc
)

if(${QT_VERSION_MAJOR} GREATER_EQUAL 6)
    qt_add_executable(BranchOfPrivates
        MANUAL_FINALIZATION
        ${PROJECT_SOURCES}
        ${APP_ICON_RESOURCE_WINDOWS}
    )
# Define target properties for Android with Qt 6 as:
#    set_property(TARGET BranchOfPrivates APPEND PROPERTY QT_ANDROID_PACKAGE_SOURCE_DIR
#                 ${CMAKE_CURRENT_SOURCE_DIR}/android)
# For more information, see https://doc.qt.io/qt-6/qt-add-executable.html#target-creation
else()
    if(ANDROID)
        add_library(BranchOfPrivates SHARED
            ${PROJECT_SOURCES}
            ${APP_ICON_RESOURCE_WINDOWS}
        )
# Define properties for Android with Qt 5 after find_package() calls as:
#    set(ANDROID_PACKAGE_SOURCE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/android")
    else()
        add_executable(BranchOfPrivates
            ${PROJECT_SOURCES}
            ${APP_ICON_RESOURCE_WINDOWS}
        )
    endif()
endif()

add_library(QXlsx SHARED IMPORTED)
set_target_properties(QXlsx PROPERTIES
    IMPORTED_IMPLIB "${CMAKE_SOURCE_DIR}/lib/msvc2019_64/QXlsxQt5.lib"
    IMPORTED_LOCATION "${CMAKE_SOURCE_DIR}/lib/msvc2019_64/QXlsxQt5.dll"
    INTERFACE_INCLUDE_DIRECTORIES "${CMAKE_SOURCE_DIR}/inc"
)
add_library(xlnt SHARED IMPORTED)
set_target_properties(xlnt PROPERTIES
    IMPORTED_IMPLIB "${CMAKE_SOURCE_DIR}/lib/msvc2019_64/xlnt.lib"
    IMPORTED_LOCATION "${CMAKE_SOURCE_DIR}/lib/msvc2019_64/xlnt.dll"
    INTERFACE_INCLUDE_DIRECTORIES "${CMAKE_SOURCE_DIR}/inc"
)
target_link_libraries(BranchOfPrivates PRIVATE Qt${QT_VERSION_MAJOR}::Widgets QXlsx xlnt)
set_target_properties(BranchOfPrivates PROPERTIES
    MACOSX_BUNDLE_GUI_IDENTIFIER my.example.com
    MACOSX_BUNDLE_BUNDLE_VERSION ${PROJECT_VERSION}
    MACOSX_BUNDLE_SHORT_VERSION_STRING ${PROJECT_VERSION_MAJOR}.${PROJECT_VERSION_MINOR}
    MACOSX_BUNDLE TRUE
    WIN32_EXECUTABLE TRUE
)

install(TARGETS BranchOfPrivates
    BUNDLE DESTINATION .
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR})

if(QT_VERSION_MAJOR EQUAL 6)
    qt_finalize_executable(BranchOfPrivates)
endif()
