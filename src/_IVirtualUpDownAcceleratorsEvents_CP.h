//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualUpDownAcceleratorsEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualUpDownAcceleratorsEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualUpDownAccelerators, EditCtlsLibU::_IVirtualUpDownAcceleratorsEvents
/// \else
///   \sa VirtualUpDownAccelerators, EditCtlsLibA::_IVirtualUpDownAcceleratorsEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualUpDownAcceleratorsEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualUpDownAcceleratorsEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualUpDownAcceleratorsEvents