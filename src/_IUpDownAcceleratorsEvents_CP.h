//////////////////////////////////////////////////////////////////////
/// \class Proxy_IUpDownAcceleratorsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IUpDownAcceleratorsEvents interface</em>
///
/// \if UNICODE
///   \sa UpDownAccelerators, EditCtlsLibU::_IUpDownAcceleratorsEvents
/// \else
///   \sa UpDownAccelerators, EditCtlsLibA::_IUpDownAcceleratorsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IUpDownAcceleratorsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IUpDownAcceleratorsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IUpDownAcceleratorsEvents