#pragma once


namespace PhishMe
{
   // COM wrapper function (ATL-like) to handle
   // auto-release of object via RAII design pattern
   template <class TYPE>
   class ComWrapper
   {
   public:

      typedef TYPE* (*CreatorFunc)();

      ComWrapper()
         : _pObject{NULL},
           _szContext(nullptr)
      {}

      ComWrapper( const char* szContext )
         : _pObject{NULL},                            // Using NULL because that's the COM method
           _szContext{szContext}
      {
         if (  szContext == nullptr )   throw runtime_error( "Missing context for ComWrapper::ComWRapper()" );
         if ( *szContext == '\0'    )   throw runtime_error( "Empty context for ComWrapper::ComWRapper()"   );
      }

      // Move constructor
      ComWrapper( ComWrapper&& wrapper )
      {
         _Move(wrapper);
      }

      virtual ~ComWrapper()
      {
         if ( _pObject != NULL )   
         {
            _pObject->Release();
            _pObject = NULL;
         }

         _szContext = nullptr;
      }

      // Move assignment operator
      ComWrapper& operator=( ComWrapper& wrapper )
      {
         _Move(wrapper);
         return *this;
      }

      // Is the contained object valid?
      bool IsValid() const
      {
         return _pObject   != NULL     &&
                _szContext != nullptr  &&
                *szContext != '\0';
      }

      //==[ GETTERS & SETTERS ]================================================
      TYPE*       GetObj()               { return _pObject;   }           
      void        SetObj( TYPE* pObj )   { _pObject = pObj; }

      const char* GetContext() const     { return _szContext; }



   protected:

      TYPE*       _pObject;
      const char* _szContext;


   private:

      void _Move( ComWrapper& wrapper )
      {
         if ( &wrapper != this )
         {
            std::swap( _pObject, wrapper._pObject );
            std::swap( _szContext, wrapper._szContext );
         }
      }
   };
}
