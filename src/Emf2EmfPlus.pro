TEMPLATE = app
TARGET = emf2emfplus
CONFIG += c++11 console utf8_sources
CONFIG -= app_bundle
CONFIG -= qt

SOURCES += \
        main.cpp

LIBS += -lgdiplus -luser32
