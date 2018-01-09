//////////////////////////////////////////////////////////////////////
/// \class Proxy_IVirtualUpDownAcceleratorEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IVirtualUpDownAcceleratorEvents interface</em>
///
/// \if UNICODE
///   \sa VirtualUpDownAccelerator, EditCtlsLibU::_IVirtualUpDownAcceleratorEvents
/// \else
///   \sa VirtualUpDownAccelerator, EditCtlsLibA::_IVirtualUpDownAcceleratorEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IVirtualUpDownAcceleratorEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IVirtualUpDownAcceleratorEvents), CComDynamicUnkArray>
{
public:
};     // Proxy_IVirtualUpDownAcceleratorEvents