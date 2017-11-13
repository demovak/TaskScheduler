//=============================================================================
// File: Test_ComWrapper.cpp - Unit tests for the ComWrapper class
//
//    Copyright (c) 2017 Jeff Reeder
//    All Rights Reserved
//=============================================================================

#include "stdafx.h"
#include "CppUnitTest.h"

#include "../ComWrapper.h"                         // Class to test

using namespace std;
using namespace Microsoft::VisualStudio::CppUnitTestFramework;
using namespace PhishMe;


namespace UnitTesting
{
   // NOTE: I don't seem to be able to write unit tests for 
   //       move constructors and move assignment operators
   TEST_CLASS(Test_ComWrapper)
   {
   public:

      // Test ComWrapper()
      TEST_METHOD(ComWrapper_DefaultConstructor)
      {
         Assert::IsFalse( ComWrapper<ITaskService>().IsValid() );
      }

      // Test ComWrapper(pointer)
      TEST_METHOD(ComWrapper_Constructor)
      {
         Assert::ExpectException<invalid_argument>( [] () { ComWrapper<ITaskService>(nullptr); } );

         // Test actual COM object
         ITaskService* pService = nullptr;
         HRESULT hr  =  CoCreateInstance( CLSID_TaskScheduler,
                                          nullptr,
                                          CLSCTX_INPROC_SERVER,
                                          IID_ITaskService,
                                          (void**) &pService );

         if ( FAILED(hr) )
         {
            Assert::Fail( L"Failed to create an instance of ITaskService" );
         }
         else
         {
            Assert::IsNotNull(pService);

            try
            {
               ComWrapper<ITaskService> service(pService);
               Assert::IsTrue( service.IsValid() );
            }
            catch ( exception& ex )
            {
               Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
            }
         }
      }

      TEST_METHOD(ComWrapper_GetObj)
      {
         ITaskService* pService = nullptr;
         HRESULT hr  =  CoCreateInstance( CLSID_TaskScheduler,
                                          nullptr,
                                          CLSCTX_INPROC_SERVER,
                                          IID_ITaskService,
                                          (void**) &pService );

         if ( FAILED(hr) )
         {
            Assert::Fail( L"Failed to create an instance of ITaskService" );
         }
         else
         {
            try
            {
               ComWrapper<ITaskService> service(pService);

               Assert::IsNotNull( service.GetObj() );
            }
            catch ( exception& ex )
            {
               Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
            }
         }
      }

      TEST_METHOD(ComWrapper_Release)
      {
         ITaskService* pService = nullptr;
         HRESULT hr  =  CoCreateInstance( CLSID_TaskScheduler,
                                          nullptr,
                                          CLSCTX_INPROC_SERVER,
                                          IID_ITaskService,
                                          (void**) &pService );

         if ( FAILED(hr) )
         {
            Assert::Fail( L"Failed to create an instance of ITaskService" );
         }
         else
         {
            try
            {
               ComWrapper<ITaskService> service(pService);
               Assert::IsTrue( service.IsValid() );

               service.Release();

               Assert::IsFalse( service.IsValid() );
               Assert::IsNull(  service.GetObj()  );
            }
            catch ( exception& ex )
            {
               Assert::Fail( EX_FN_MSG( ex, "Exception thrown" ) );
            }
         }
      }
   };
}