#pragma once


namespace PhishMe
{
   class com_exception  :  public std::exception
   {
   public:

      com_exception( const char* message, 
                     HRESULT     hr = -1L );

      virtual ~com_exception();

      virtual const char* what() const;

   private:

      std::string _message;

      com_exception() = default;                 // Require use of custom constructor
   };
}

