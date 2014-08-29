#ifndef PLANTSCHEDULER_H
#define PLANTSCHEDULER_H

typedef struct _plant
{
	int plant_type;
	std::string common_name;
	std::string bio_name;
	int pot_size;
	int install_height;
	int install_spacing;
}plant;

#endif