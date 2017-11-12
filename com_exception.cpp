#include "stdafx.h"

#include "com_exception.h"

using namespace std;
using namespace PhishMe;


com_exception::com_exception( const char* message,
                              HRESULT     hr /* = -1L */ )
{
   if ( hr < 0 )
   {
      _message = string(message);
   }
   else
   {
      _message  =  string(message)  +  ".  Result: "  +  to_string(hr);
   }
}


com_exception::~com_exception()
{
}


const char* com_exception::what() const
{
   return _message.c_str();
}
