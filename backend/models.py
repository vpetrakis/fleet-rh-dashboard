from sqlalchemy import Column, Integer, String, ForeignKey
from sqlalchemy.orm import relationship
from database import Base

class Vessel(Base):
    __tablename__ = "vessels"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, unique=True, index=True, nullable=False) # e.g. MV ALEXIS [cite: 72]
    main_engine_components = relationship("MainEngineComponent", back_populates="vessel")

class MainEngineComponent(Base):
    __tablename__ = "main_engine_components"
    id = Column(Integer, primary_key=True, index=True)
    vessel_id = Column(Integer, ForeignKey("vessels.id"))
    component_name = Column(String) # e.g. CYLINDER COVER [cite: 73]
    cylinder_number = Column(Integer) # e.g. 1 [cite: 73]
    running_hours = Column(Integer) # e.g. 71225 [cite: 72]
    vessel = relationship("Vessel", back_populates="main_engine_components")