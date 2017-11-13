//=============================================================================
// File: Test_com_exception.cpp - Unit tests for the com_exception class
//
//    Copyright (c) 2017 Jeff Reeder
//    All Rights Reserved
//=============================================================================

#include "stdafx.h"

#include "CppUnitTest.h"

#include "../com_exception.h"                      // Class to test

using namespace std;
using namespace Microsoft::VisualStudio::CppUnitTestFramework;
using namespace PhishMe;


namespace UnitTesting
{		
	TEST_CLASS(Test_com_exception)
	{
	public:
		
      // Test com_exception( const char* )
      TEST_METHOD(com_exception_DefaultConstructor)
      {
         // Verify 
         Assert::ExpectException<invalid_argument>( []() { com_exception(nullptr); } );
         Assert::ExpectException<invalid_argument>( []() { com_exception( "" );    } );
      }

      // Test com_exception( const char*, HRESULT )
      TEST_METHOD(com_exception_HresultConstructor)
		{
         // Verify 
         Assert::ExpectException<invalid_argument>( [] () { com_exception(nullptr);       } );
         Assert::ExpectException<invalid_argument>( [] () { com_exception("");            } );
         Assert::ExpectException<invalid_argument>( [] () { com_exception("Hello", -1L ); } );
      }

      // Test what()
      TEST_METHOD(com_exception_What)
      {
         // Verify what() content
         Assert::AreEqual( string( "Hello" ).c_str(),
                           com_exception( "Hello" ).what() );

         Assert::AreEqual( string( "Hello.  Result: 123" ).c_str(),
                           com_exception( "Hello", 123L ).what() );
      }
   };
}